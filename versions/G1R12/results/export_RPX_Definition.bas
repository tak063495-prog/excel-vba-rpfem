Option Explicit
Public Type RPX_Region
    materialId As Long
    rigidId As Long
    pointCount As Long
    points() As Long
    area As Double
End Type
Public Type RPX_FaceCondition
    regionId As Long
    startPoint As Long
    endPoint As Long
    wrap As Boolean
    kind As String
    rigidId As Long
    load0(0 To 1) As Double
    load1(0 To 1) As Double
End Type
Public DPCount As Long, DRCount As Long, DFCount As Long
Public DPXY() As Double, DRegions() As RPX_Region, DFaces() As RPX_FaceCondition
Public DSegments As Object, DTarget As Long, DSpacing As Double, DTolerance As Double
Private Function DefinitionCross(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Double
    DefinitionCross = (DPXY(b, 0) - DPXY(a, 0)) * (DPXY(c, 1) - DPXY(a, 1)) - (DPXY(b, 1) - DPXY(a, 1)) * (DPXY(c, 0) - DPXY(a, 0))
End Function
Private Function OnRegionBoundary(ByVal point As Long, ByVal region As Long) As Boolean
    Dim i As Long, a As Long, b As Long
    For i = 0 To DRegions(region).pointCount - 1
        a = DRegions(region).points(i): b = DRegions(region).points((i + 1) Mod DRegions(region).pointCount)
        If RPX_OnSegment(DPXY(point, 0), DPXY(point, 1), DPXY(a, 0), DPXY(a, 1), DPXY(b, 0), DPXY(b, 1)) Then OnRegionBoundary = True: Exit Function
    Next i
End Function
Private Sub ValidateRegions()
    Dim keys As Variant, p As Variant, q As Variant, i As Long, j As Long, r As Long, s As Long, id As Long
    Dim seen As Object, inside As Boolean, outside As Boolean, allBoundary As Boolean
    keys = DSegments.keys
    For i = 0 To DSegments.count - 1
        p = DSegments(keys(i))
        For j = 0 To i - 1
            q = DSegments(keys(j))
            If DefinitionCross(p(0), p(1), q(0)) * DefinitionCross(p(0), p(1), q(1)) < -1E-20 And DefinitionCross(q(0), q(1), p(0)) * DefinitionCross(q(0), q(1), p(1)) < -1E-20 Then err.Raise 5, , "領域境界が定義点以外で交差しています。"
        Next j
    Next i
    For r = 0 To DRCount - 1
        Set seen = RPX_Row()
        For i = 0 To DRegions(r).pointCount - 1
            id = DRegions(r).points(i)
            If seen.Exists(id) Then err.Raise 5, , "領域境界が同じ点を複数回通っています。"
            seen(id) = True
        Next i
        For s = 0 To DRCount - 1
            If s <> r Then
                inside = False: outside = False: allBoundary = True
                For i = 0 To DRegions(r).pointCount - 1
                    id = DRegions(r).points(i)
                    If Not OnRegionBoundary(id, s) Then
                        allBoundary = False
                        If RPX_InRegion(DPXY(id, 0), DPXY(id, 1), s) Then inside = True Else outside = True
                    End If
                Next i
                If inside And outside Then err.Raise 5, , "領域が部分的に重複しています。隣接領域または内包領域として定義してください。"
                If allBoundary And Abs(DRegions(r).area - DRegions(s).area) < DTolerance Then err.Raise 5, , "領域が重複しています。"
            End If
        Next s
    Next r
    For i = 0 To DPCount - 1
        If RPX_RegionAt(DPXY(i, 0), DPXY(i, 1)) < 0 Then err.Raise 5, , "どの領域にも属さない定義点があります。"
    Next i
End Sub
Public Function RPX_SegmentDistance(ByVal px As Double, ByVal py As Double, ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Double
    Dim dx As Double, dy As Double, t As Double
    dx = bx - ax: dy = by - ay
    If dx * dx + dy * dy <= 0# Then err.Raise 5, , "ZERO_LENGTH_SEGMENT"
    t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    RPX_SegmentDistance = Sqr((px - ax - t * dx) ^ 2 + (py - ay - t * dy) ^ 2)
End Function
Public Function RPX_UnionArea() As Double
    Dim i As Long, j As Long, k As Long, contained As Boolean, allInside As Boolean
    For i = 0 To DRCount - 1
        contained = False
        For j = 0 To DRCount - 1
            If DRegions(j).area > DRegions(i).area + 0.0000000001 Then
                allInside = True
                For k = 0 To DRegions(i).pointCount - 1
                    If Not RPX_InRegion(DPXY(DRegions(i).points(k), 0), DPXY(DRegions(i).points(k), 1), j) Then allInside = False: Exit For
                Next k
                If allInside Then contained = True: Exit For
            End If
        Next j
        If Not contained Then RPX_UnionArea = RPX_UnionArea + DRegions(i).area
    Next i
End Function

Public Function RPX_OnSegment(ByVal px As Double, ByVal py As Double, ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Boolean
    Dim dx As Double, dy As Double, length2 As Double, s As Double, cross As Double
    dx = bx - ax: dy = by - ay: length2 = dx * dx + dy * dy
    If length2 <= DTolerance * DTolerance Then Exit Function
    s = ((px - ax) * dx + (py - ay) * dy) / length2
    cross = Abs((px - ax) * dy - (py - ay) * dx) / Sqr(length2)
    RPX_OnSegment = (s >= -DTolerance And s <= 1# + DTolerance And cross <= DTolerance)
End Function
Public Function RPX_InRegion(ByVal x As Double, ByVal y As Double, ByVal regionId As Long) As Boolean
    Dim i As Long, j As Long, a As Long, b As Long, insideFlag As Boolean, xx As Double
    With DRegions(regionId)
        j = .pointCount - 1
        For i = 0 To .pointCount - 1
            a = .points(j): b = .points(i)
            If RPX_OnSegment(x, y, DPXY(a, 0), DPXY(a, 1), DPXY(b, 0), DPXY(b, 1)) Then RPX_InRegion = True: Exit Function
            If (DPXY(a, 1) > y) <> (DPXY(b, 1) > y) Then
                xx = DPXY(a, 0) + (y - DPXY(a, 1)) * (DPXY(b, 0) - DPXY(a, 0)) / (DPXY(b, 1) - DPXY(a, 1))
                If x < xx Then insideFlag = Not insideFlag
            End If
            j = i
        Next i
    End With
    RPX_InRegion = insideFlag
End Function
Public Function RPX_RegionAt(ByVal x As Double, ByVal y As Double) As Long
    Dim r As Long, smallest As Double
    RPX_RegionAt = -1: smallest = 1E+300
    For r = 0 To DRCount - 1
        If RPX_InRegion(x, y, r) Then
            If DRegions(r).area < smallest Then smallest = DRegions(r).area: RPX_RegionAt = r
        End If
    Next r
End Function
Public Sub RPX_PrepareDefinition()
    Dim r As Long, i As Long, j As Long, k As Long, a As Long, b As Long, tmp As Long, count As Long
    Dim points() As Long, coordinate() As Double, val As Double, area As Double, dx As Double, dy As Double, key As String, expanded As Collection
    Set DSegments = RPX_Row(): DTolerance = 0.00000001
    For r = 0 To DRCount - 1
        area = 0#: Set expanded = New Collection
        With DRegions(r)
            If .pointCount < 3 Then err.Raise 5, , "REGION_NEEDS_THREE_POINTS"
            For i = 0 To .pointCount - 1
                a = .points(i): b = .points((i + 1) Mod .pointCount)
                If a < 0 Or b < 0 Or a >= DPCount Or b >= DPCount Or a = b Then err.Raise 5, , "INVALID_REGION_POINT"
                area = area + DPXY(a, 0) * DPXY(b, 1) - DPXY(b, 0) * DPXY(a, 1)
                ' Split every straight segment at all intervening definition points.
                ReDim points(0 To DPCount - 1): ReDim coordinate(0 To DPCount - 1): count = 0
                dx = DPXY(b, 0) - DPXY(a, 0): dy = DPXY(b, 1) - DPXY(a, 1)
                For j = 0 To DPCount - 1
                    If RPX_OnSegment(DPXY(j, 0), DPXY(j, 1), DPXY(a, 0), DPXY(a, 1), DPXY(b, 0), DPXY(b, 1)) Then
                        points(count) = j: coordinate(count) = (DPXY(j, 0) - DPXY(a, 0)) * dx + (DPXY(j, 1) - DPXY(a, 1)) * dy: count = count + 1
                    End If
                Next j
                For j = 1 To count - 1
                    tmp = points(j): val = coordinate(j): k = j - 1
                    Do While k >= 0
                        If coordinate(k) <= val Then Exit Do
                        points(k + 1) = points(k): coordinate(k + 1) = coordinate(k): k = k - 1
                    Loop
                    points(k + 1) = tmp: coordinate(k + 1) = val
                Next j
                For j = 0 To count - 2
                    expanded.Add points(j)
                    key = RPX_EdgeKey(points(j), points(j + 1)): DSegments(key) = Array(points(j), points(j + 1))
                Next j
            Next i
            .pointCount = expanded.count: ReDim .points(0 To .pointCount - 1)
            For i = 0 To .pointCount - 1: .points(i) = CLng(expanded(i + 1)): Next i
            .area = Abs(area) / 2#: If .area <= 0.000000000001 Then err.Raise 5, , "ZERO_REGION_AREA"
        End With
    Next r
    ' Face paths follow the selected region's boundary order, including bends.
    Call ValidateRegions
    For i = 0 To DFCount - 1
        r = DFaces(i).regionId: a = -1: b = -1
        If r < 0 Or r >= DRCount Then err.Raise 5, , "INVALID_FACE_REGION"
        For j = 0 To DRegions(r).pointCount - 1
            If DRegions(r).points(j) = DFaces(i).startPoint Then a = j
            If DRegions(r).points(j) = DFaces(i).endPoint Then b = j
        Next j
        If a < 0 Or b < 0 Or a = b Then err.Raise 5, , "FACE_ENDPOINT_NOT_ON_REGION_LOOP"
        If b < a And Not DFaces(i).wrap Then err.Raise 5, , "FACE_PATH_REQUIRES_EXPLICIT_WRAP"
        If b > a And DFaces(i).wrap Then err.Raise 5, , "循環指定の経路が領域の末尾をまたいでいません。"
    Next i
End Sub

Public Function RPX_FaceContains(ByVal faceId As Long, ByVal px As Double, ByVal py As Double) As Boolean
    Dim r As Long, i As Long, a As Long, b As Long, active As Boolean, stepCount As Long
    r = DFaces(faceId).regionId
    For i = 0 To DRegions(r).pointCount - 1
        If DRegions(r).points(i) = DFaces(faceId).startPoint Then Exit For
    Next i
    For stepCount = 1 To DRegions(r).pointCount
        a = DRegions(r).points(i): b = DRegions(r).points((i + 1) Mod DRegions(r).pointCount)
        If RPX_OnSegment(px, py, DPXY(a, 0), DPXY(a, 1), DPXY(b, 0), DPXY(b, 1)) Then RPX_FaceContains = True: Exit Function
        If b = DFaces(faceId).endPoint Then Exit For
        i = (i + 1) Mod DRegions(r).pointCount
    Next stepCount
End Function
Public Sub RPX_MapFaces()
    Dim id As Long, faceId As Long, j As Long, px As Double, py As Double, fixedX As Boolean, fixedY As Boolean
    Set RBoundary = RPX_Row()
    For id = 0 To RNEdge - 1
        With REdges(id)
            px = (RXY(.a, 0) + RXY(.b, 0)) / 2#: py = (RXY(.a, 1) + RXY(.b, 1)) / 2#
            fixedX = False: fixedY = False: .kind = "load": .rigidId = -1
            For j = 0 To 1: .load0(j) = 0#: .load1(j) = 0#: Next j
            For faceId = 0 To DFCount - 1
                If RPX_FaceContains(faceId, px, py) Then
                    If .element(1) >= 0 Then err.Raise 5, , "FACE_CONDITION_MUST_BE_ON_EXPOSED_BOUNDARY"
                    For j = 0 To 1: .load0(j) = .load0(j) + DFaces(faceId).load0(j): .load1(j) = .load1(j) + DFaces(faceId).load1(j): Next j
                    Select Case DFaces(faceId).kind
                        Case "fixed": fixedX = True: fixedY = True
                        Case "roller_x": fixedX = True
                        Case "roller_y": fixedY = True
                        Case "rigid"
                            If .rigidId >= 0 And .rigidId <> DFaces(faceId).rigidId Then err.Raise 5, , "CONFLICTING_RIGID_FACE"
                            .rigidId = DFaces(faceId).rigidId: .kind = "rigid"
                        Case "load"
                        Case Else: err.Raise 5, , "UNKNOWN_FACE_RESTRAINT"
                    End Select
                End If
            Next faceId
            If .kind = "rigid" And (fixedX Or fixedY) Then err.Raise 5, , "RIGID_AND_FIXED_FACE_CONFLICT"
            If fixedX And fixedY Then
                .kind = "fixed"
            ElseIf fixedX Then
                .kind = "roller_x"
            ElseIf fixedY Then
                .kind = "roller_y"
            End If
            If .element(1) < 0 Then RBoundary(RPX_EdgeKey(.a, .b)) = Array(.kind, .rigidId, .load0(0), .load0(1), .load1(0), .load1(1))
        End With
    Next id
End Sub