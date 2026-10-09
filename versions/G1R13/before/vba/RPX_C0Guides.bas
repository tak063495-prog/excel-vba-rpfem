Attribute VB_Name = "RPX_C0Guides"
Option Explicit
' Slope-parallel guides for c=0 soil only.
' A c>0 region does not receive bands. A band that would leave the c=0 soil,
' or enter a rigid body, is cut off. Guides do not add faces, materials, loads, or rigids.
' The automatic Fs path (RPX_C0Select) turns guides on for one candidate at a time.
' C0_UPPER_ADAPT=1 still injects guides on a normal mesh generation.
' Ratios stay at or above 0.08 of the face rise. Thinner bands made the lower bound fail.
' BAND is the default. CHORD is not used by the automatic trial.
Public C0ForceGuides As Boolean
Public C0ForceSpacing As Double
Public C0ForceRatios As String
Public C0ForceRoute As String
Public RPX_C0Ran As Boolean
Public RPX_C0Adopted As String
Public RPX_C0TrialLog As String
Public RPX_C0SplitMesh As Boolean
Public RPX_C0LowerElements As Long
Public RPX_C0UpperElements As Long

Public Function RPX_C0HasZeroCohesion() As Boolean
    ' Used soils only. An unused material row must not change the route.
    Dim i As Long, id As Long, seen As Long
    If RM <= 0 Then Exit Function
    If DRCount > 0 Then
        For i = 0 To DRCount - 1
            If DRegions(i).rigidId < 0 Then
                id = DRegions(i).materialId
                If id >= 0 And id < RM Then
                    seen = seen + 1
                    If Abs(RMaterials(id).cohesion) <= 0.000000000001 Then
                        RPX_C0HasZeroCohesion = True
                        Exit Function
                    End If
                End If
            End If
        Next i
        If seen > 0 Then Exit Function
    End If
    If re > 0 Then
        For i = 0 To re - 1
            If RRigidId(i) < 0 Then
                id = RMatId(i)
                If id >= 0 And id < RM Then
                    If Abs(RMaterials(id).cohesion) <= 0.000000000001 Then
                        RPX_C0HasZeroCohesion = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
End Function

Public Function RPX_C0Problem() As Boolean
    ' Calibrated band trial. Fixed Fs probes from the single-soil associated case.
    If RM <> 1 Or RR <> 0 Or DRCount <> 1 Then Exit Function
    If Abs(RMaterials(0).cohesion) > 0.000000000001 Then Exit Function
    If RMaterials(0).friction <= 0# Or RMaterials(0).friction >= 89# Then Exit Function
    If Abs(RMaterials(0).dilation - RMaterials(0).friction) > 0.0001 Then Exit Function
    RPX_C0Problem = True
End Function

Public Function RPX_C0UpperAdaptOn() As Boolean
    If C0ForceGuides Then
        If RPX_C0HasZeroCohesion() Then RPX_C0UpperAdaptOn = True
        Exit Function
    End If
    If CLng(RPX_Setting("C0_UPPER_ADAPT", 0)) = 0 Then Exit Function
    If Not RPX_C0Problem() Then Exit Function
    RPX_C0UpperAdaptOn = True
End Function

Public Function RPX_C0GuidePolylines(ByVal alongSpacing As Double) As Collection
    Dim route As String, ratioText As String, parts As Variant, ratios() As Double, nRatio As Long, i As Long
    Dim guides As New Collection
    Set RPX_C0GuidePolylines = guides
    If alongSpacing <= 0.000000001 Then Exit Function
    If Not RPX_C0UpperAdaptOn() Then Exit Function
    If Len(C0ForceRoute) > 0 Then
        route = UCase$(Trim$(C0ForceRoute))
    Else
        route = UCase$(Trim$(CStr(RPX_Setting("C0_MESH_ROUTE", "BAND"))))
    End If
    If route <> "CHORD" Then route = "BAND"
    If Len(C0ForceRatios) > 0 Then
        ratioText = Replace$(C0ForceRatios, " ", "")
    Else
        ratioText = Replace$(CStr(RPX_Setting("C0_LAYER_DEPTH_RATIOS", "0.08,0.15")), " ", "")
    End If
    parts = Split(ratioText, ",")
    nRatio = 0
    ReDim ratios(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If IsNumeric(parts(i)) Then
                ratios(nRatio) = CDbl(parts(i))
                nRatio = nRatio + 1
            End If
        End If
    Next i
    If nRatio = 0 Then Exit Function
    Call AppendGuides(guides, route, ratios, nRatio, alongSpacing)
End Function

Private Sub AppendGuides(ByRef guides As Collection, ByVal route As String, ByRef ratios() As Double, ByVal nRatio As Long, ByVal alongSpacing As Double)
    Dim regionId As Long
    For regionId = 0 To DRCount - 1
        If RegionIsC0(regionId) Then Call AppendRegionGuides(guides, regionId, route, ratios, nRatio, alongSpacing)
    Next regionId
End Sub

Private Sub AppendRegionGuides(ByRef guides As Collection, ByVal regionId As Long, ByVal route As String, ByRef ratios() As Double, ByVal nRatio As Long, ByVal alongSpacing As Double)
    Dim nLoop As Long, i As Long, a As Long, b As Long, j As Long
    Dim mx As Double, my As Double, covered As Boolean, faceId As Long
    Dim dx As Double, dy As Double, length As Double, angle As Double, rise As Double
    Dim clockwise As Boolean, area2 As Double, nx As Double, ny As Double
    Dim depth As Double, p0x As Double, p0y As Double, p1x As Double, p1y As Double
    Dim line As Collection
    nLoop = DRegions(regionId).pointCount
    If nLoop < 3 Then Exit Sub
    area2 = 0#
    For i = 0 To nLoop - 1
        a = DRegions(regionId).points(i): b = DRegions(regionId).points((i + 1) Mod nLoop)
        area2 = area2 + DPXY(a, 0) * DPXY(b, 1) - DPXY(b, 0) * DPXY(a, 1)
    Next i
    clockwise = (area2 < 0#)
    For i = 0 To nLoop - 1
        a = DRegions(regionId).points(i): b = DRegions(regionId).points((i + 1) Mod nLoop)
        If EdgeIsShared(regionId, a, b) Then GoTo nextEdge
        mx = 0.5 * (DPXY(a, 0) + DPXY(b, 0)): my = 0.5 * (DPXY(a, 1) + DPXY(b, 1))
        covered = False
        For faceId = 0 To DFCount - 1
            If DFaces(faceId).kind <> "load" Then
                If RPX_FaceContains(faceId, mx, my) Then covered = True: Exit For
            End If
        Next faceId
        If covered Then GoTo nextEdge
        dx = DPXY(b, 0) - DPXY(a, 0): dy = DPXY(b, 1) - DPXY(a, 1)
        length = Sqr(dx * dx + dy * dy)
        If length <= 0.00000001 Then GoTo nextEdge
        If Abs(dx) <= 0.000000000001 Then
            angle = 90#
        Else
            angle = Abs(Atn(dy / dx) * 45# / Atn(1#))
        End If
        If angle < 8# Or angle > 82# Then GoTo nextEdge
        If clockwise Then
            nx = dy / length: ny = -dx / length
        Else
            nx = -dy / length: ny = dx / length
        End If
        rise = Abs(dy)
        If rise <= 0.00000001 Then rise = length
        For j = 0 To nRatio - 1
            depth = ratios(j) * rise
            If depth > 0.000001 Then
                p0x = DPXY(a, 0) + depth * nx: p0y = DPXY(a, 1) + depth * ny
                p1x = DPXY(b, 0) + depth * nx: p1y = DPXY(b, 1) + depth * ny
                If route = "CHORD" Then
                    Set line = ChordSamples(p0x, p0y, dx, dy, alongSpacing)
                Else
                    Set line = SegmentSamples(p0x, p0y, p1x, p1y, alongSpacing)
                End If
                Call KeepInsideC0(guides, line)
            End If
        Next j
nextEdge:
    Next i
End Sub

Private Function RegionIsC0(ByVal regionId As Long) As Boolean
    Dim mat As Long
    If regionId < 0 Or regionId >= DRCount Then Exit Function
    If DRegions(regionId).rigidId >= 0 Then Exit Function
    mat = DRegions(regionId).materialId
    If mat < 0 Or mat >= RM Then Exit Function
    If Abs(RMaterials(mat).cohesion) <= 0.000000000001 Then RegionIsC0 = True
End Function

Private Function C0PointInside(ByVal x As Double, ByVal y As Double) As Boolean
    C0PointInside = RegionIsC0(RPX_RegionAt(x, y))
End Function

Private Sub KeepInsideC0(ByRef guides As Collection, ByRef line As Collection)
    Dim i As Long, x As Double, y As Double
    Dim run As Collection, pt As Variant
    If line Is Nothing Then Exit Sub
    Set run = New Collection
    For i = 1 To line.count
        pt = line(i)
        x = CDbl(pt(0)): y = CDbl(pt(1))
        If C0PointInside(x, y) Then
            run.Add Array(x, y)
        Else
            If run.count >= 2 Then guides.Add run
            Set run = New Collection
        End If
    Next i
    If run.count >= 2 Then guides.Add run
End Sub

Private Function EdgeIsShared(ByVal regionId As Long, ByVal a As Long, ByVal b As Long) As Boolean
    Dim r As Long, i As Long, p As Long, q As Long, nLoop As Long
    For r = 0 To DRCount - 1
        If r <> regionId Then
            nLoop = DRegions(r).pointCount
            For i = 0 To nLoop - 1
                p = DRegions(r).points(i): q = DRegions(r).points((i + 1) Mod nLoop)
                If (p = a And q = b) Or (p = b And q = a) Then EdgeIsShared = True: Exit Function
            Next i
        End If
    Next r
End Function

Private Function SegmentSamples(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal spacing As Double) As Collection
    Dim span As Double, divisions As Long, j As Long, t As Double, line As New Collection
    span = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)
    divisions = -Int(-span / spacing)
    If divisions < 1 Then divisions = 1
    For j = 0 To divisions
        t = j / divisions
        line.Add Array(x0 + t * (x1 - x0), y0 + t * (y1 - y0))
    Next j
    Set SegmentSamples = line
End Function

Private Function ChordSamples(ByVal ox As Double, ByVal oy As Double, ByVal dx As Double, ByVal dy As Double, ByVal spacing As Double) As Collection
    Dim i As Long, a As Long, b As Long, ex As Double, ey As Double, det As Double, t As Double, s As Double
    Dim nLoop As Long, tMin As Double, tMax As Double, hits As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    nLoop = DRegions(0).pointCount
    tMin = 0#: tMax = 0#: hits = 0
    For i = 0 To nLoop - 1
        a = DRegions(0).points(i): b = DRegions(0).points((i + 1) Mod nLoop)
        ex = DPXY(b, 0) - DPXY(a, 0): ey = DPXY(b, 1) - DPXY(a, 1)
        det = dx * ey - dy * ex
        If Abs(det) >= 0.00000000000001 Then
            t = ((DPXY(a, 0) - ox) * ey - (DPXY(a, 1) - oy) * ex) / det
            s = ((DPXY(a, 0) - ox) * dy - (DPXY(a, 1) - oy) * dx) / det
            If s >= -0.000000001 And s <= 1.000000001 Then
                If hits = 0 Then
                    tMin = t: tMax = t
                Else
                    If t < tMin Then tMin = t
                    If t > tMax Then tMax = t
                End If
                hits = hits + 1
            End If
        End If
    Next i
    If hits < 2 Or tMax - tMin < 0.000001 Then Exit Function
    x0 = ox + tMin * dx: y0 = oy + tMin * dy
    x1 = ox + tMax * dx: y1 = oy + tMax * dy
    Set ChordSamples = SegmentSamples(x0, y0, x1, y1, spacing)
End Function
