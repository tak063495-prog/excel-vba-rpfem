Attribute VB_Name = "RPX_FlowMesh"
Option Explicit
' Pure array geometry candidate. No global model, field, load, Fs or root access.
' Candidate connectivity alone is never a physical certificate or a saved field.
Private Type FMMetric
    xx As Double
    xy As Double
    yy As Double
End Type
Private Type FMEdge
    a As Long
    b As Long
    e As Long
    f As Long
End Type
Private Type FMChoice
    e As Long
    f As Long
    a As Long
    b As Long
    c As Long
    d As Long
    p0 As Long
    p1 As Long
    p2 As Long
    q0 As Long
    q1 As Long
    q2 As Long
    ordinal As Long
    gain As Double
End Type
Private Type FMFloat
    value As Double
End Type
Private Type FMBits
    low As Long
    high As Long
End Type

Private Function FMFinite(ByVal value As Double) As Boolean
    Dim raw As FMFloat, bits As FMBits
    raw.value = value: LSet bits = raw
    FMFinite = ((bits.high And &H7FF00000) <> &H7FF00000)
End Function
Private Function FMKey(ByVal a As Long, ByVal b As Long) As String
    If a < b Then FMKey = CStr(a) & ":" & CStr(b) Else FMKey = CStr(b) & ":" & CStr(a)
End Function
Private Function FMDet(ByRef points() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long) As Double
    FMDet = (points(b, 0) - points(a, 0)) * (points(c, 1) - points(a, 1)) - (points(b, 1) - points(a, 1)) * (points(c, 0) - points(a, 0))
End Function
Private Sub FMOrient(ByRef points() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long, ByRef u As Long, ByRef v As Long, ByRef w As Long)
    u = a
    If FMDet(points, a, b, c) > 0# Then v = b: w = c Else v = c: w = b
End Sub
Private Function FMEdges(ByRef cells() As Long, ByRef edges() As FMEdge, ByRef count As Long) As Object
    Dim lookup As Object, e As Long, k As Long, a As Long, b As Long, index As Long, key As String
    Set lookup = CreateObject("Scripting.Dictionary")
    ReDim edges(0 To 3 * (UBound(cells, 1) + 1) - 1): count = 0
    For e = 0 To UBound(cells, 1)
        For k = 0 To 2
            a = cells(e, k): b = cells(e, (k + 1) Mod 3): key = FMKey(a, b)
            If lookup.Exists(key) Then
                index = CLng(lookup(key))
                If edges(index).f <> -1 Then err.Raise 5, "RPX_FlowMesh", "FLOW_NONMANIFOLD_EDGE"
                edges(index).f = e
            Else
                lookup.Add key, count
                If a < b Then edges(count).a = a: edges(count).b = b Else edges(count).a = b: edges(count).b = a
                edges(count).e = e: edges(count).f = -1: count = count + 1
            End If
        Next k
    Next e
    Set FMEdges = lookup
End Function
Private Function FMOpposite(ByRef cells() As Long, ByVal e As Long, ByVal a As Long, ByVal b As Long) As Long
    Dim k As Long
    For k = 0 To 2
        If cells(e, k) <> a And cells(e, k) <> b Then FMOpposite = cells(e, k): Exit Function
    Next k
    err.Raise 5, "RPX_FlowMesh", "FLOW_INVALID_EDGE_CAVITY"
End Function
Private Function FMMin(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then FMMin = a Else FMMin = b
End Function
Private Function FMQuadratic(ByVal x As Double, ByVal y As Double, ByRef metric As FMMetric) As Double
    FMQuadratic = (x * metric.xx + y * metric.xy) * x + (x * metric.xy + y * metric.yy) * y
End Function
Private Function FMQuality(ByRef points() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long, ByRef metric As FMMetric) As Double
    Dim determinant As Double, denominator As Double, md As Double, value As Double
    determinant = FMDet(points, a, b, c)
    If determinant <= 0# Then FMQuality = -1#: Exit Function
    denominator = FMQuadratic(points(b, 0) - points(a, 0), points(b, 1) - points(a, 1), metric)
    denominator = denominator + FMQuadratic(points(c, 0) - points(b, 0), points(c, 1) - points(b, 1), metric)
    denominator = denominator + FMQuadratic(points(a, 0) - points(c, 0), points(a, 1) - points(c, 1), metric)
    md = metric.xx * metric.yy - metric.xy * metric.xy
    If denominator <= 0# Or md <= 0# Or Not FMFinite(denominator) Or Not FMFinite(md) Then err.Raise 5, "RPX_FlowMesh", "FLOW_INVALID_QUALITY_METRIC"
    value = 2# * Sqr(3#) * determinant * Sqr(md) / denominator
    If Not FMFinite(value) Then err.Raise 5, "RPX_FlowMesh", "FLOW_NONFINITE_QUALITY"
    FMQuality = value
End Function
Private Sub FMSourceMetrics(ByRef points() As Double, ByRef cells() As Long, ByRef velocity() As Double, ByRef metrics() As FMMetric)
    Dim e As Long, k As Long, j As Long, nextNode As Long, otherNode As Long
    Dim g(0 To 2, 0 To 1) As Double, p2(0 To 5, 0 To 1) As Double, jac(0 To 1, 0 To 1) As Double
    Dim det As Double, third As Double, a As Double, b As Double, c As Double, largest As Double, disc As Double
    Dim high As Double, low As Double, floor As Double, alpha As Double, beta As Double, metricScale As Double
    ReDim metrics(0 To UBound(cells, 1)): third = 1# / 3#
    For e = 0 To UBound(cells, 1)
        det = FMDet(points, cells(e, 0), cells(e, 1), cells(e, 2))
        For k = 0 To 2
            nextNode = cells(e, (k + 1) Mod 3): otherNode = cells(e, (k + 2) Mod 3)
            g(k, 0) = (points(nextNode, 1) - points(otherNode, 1)) / det
            g(k, 1) = (points(otherNode, 0) - points(nextNode, 0)) / det
        Next k
        For k = 0 To 2
            For j = 0 To 1
                p2(k, j) = (4# * third - 1#) * g(k, j)
                p2(3 + k, j) = 4# * (third * g((k + 1) Mod 3, j) + third * g(k, j))
            Next j
        Next k
        jac(0, 0) = 0#: jac(0, 1) = 0#: jac(1, 0) = 0#: jac(1, 1) = 0#
        For k = 0 To 5
            For j = 0 To 1
                jac(j, 0) = jac(j, 0) + velocity(12 * e + 2 * k + j) * p2(k, 0)
                jac(j, 1) = jac(j, 1) + velocity(12 * e + 2 * k + j) * p2(k, 1)
            Next j
        Next k
        a = jac(0, 0) * jac(0, 0) + jac(1, 0) * jac(1, 0)
        b = jac(0, 0) * jac(0, 1) + jac(1, 0) * jac(1, 1)
        c = jac(0, 1) * jac(0, 1) + jac(1, 1) * jac(1, 1)
        If Not FMFinite(a) Or Not FMFinite(b) Or Not FMFinite(c) Then err.Raise 5, "RPX_FlowMesh", "FLOW_NONFINITE_SOURCE_METRIC"
        largest = Abs(a): If Abs(b) > largest Then largest = Abs(b)
        If Abs(c) > largest Then largest = Abs(c)
        disc = 0#
        If largest > 0# Then disc = largest * Sqr(((a - c) / largest) ^ 2 + 4# * (b / largest) ^ 2)
        high = (a + c) / 2# + disc / 2#: low = (a + c) / 2# - disc / 2#
        metricScale = high: If metricScale < 0.00000000000000000001 Then metricScale = 0.00000000000000000001
        floor = metricScale / 9#
        If low >= floor Then
            metrics(e).xx = a / metricScale: metrics(e).xy = b / metricScale: metrics(e).yy = c / metricScale
        ElseIf high <= floor Then
            metrics(e).xx = 1# / 9#: metrics(e).xy = 0#: metrics(e).yy = 1# / 9#
        Else
            ' Spectral floor via a polynomial in the symmetric 2x2 matrix.
            ' This is the same eigenvalue clamp as (vec*max(val,hi/9))*vec.T/hi.
            alpha = (high - floor) / (high - low): beta = floor - alpha * low
            metrics(e).xx = (alpha * a + beta) / metricScale
            metrics(e).xy = alpha * b / metricScale
            metrics(e).yy = (alpha * c + beta) / metricScale
        End If
    Next e
End Sub
Private Function FMPrecedes(ByRef choices() As FMChoice, ByVal a As Long, ByVal b As Long) As Boolean
    If choices(a).gain = choices(b).gain Then
        FMPrecedes = (choices(a).ordinal < choices(b).ordinal)
    Else
        FMPrecedes = (choices(a).gain > choices(b).gain)
    End If
End Function
Private Sub FMSort(ByRef order() As Long, ByRef choices() As FMChoice, ByVal first As Long, ByVal last As Long)
    Dim left As Long, right As Long, pivot As Long, temporary As Long
    left = first: right = last: pivot = order((first + last) \ 2)
    Do While left <= right
        Do While FMPrecedes(choices, order(left), pivot): left = left + 1: Loop
        Do While FMPrecedes(choices, pivot, order(right)): right = right - 1: Loop
        If left <= right Then
            temporary = order(left): order(left) = order(right): order(right) = temporary
            left = left + 1: right = right - 1
        End If
    Loop
    If first < right Then FMSort order, choices, first, right
    If left < last Then FMSort order, choices, left, last
End Sub
Private Function FMAreas(ByRef points() As Double, ByRef cells() As Long, ByRef materialId() As Long) As Object
    Dim areas As Object, e As Long, key As String, det As Double
    Set areas = CreateObject("Scripting.Dictionary")
    For e = 0 To UBound(cells, 1)
        det = FMDet(points, cells(e, 0), cells(e, 1), cells(e, 2))
        If Not FMFinite(det) Or det <= 0.000000000001 Then err.Raise 5, "RPX_FlowMesh", "FLOW_NONPOSITIVE_TRIANGLE"
        key = CStr(materialId(e))
        If Not areas.Exists(key) Then areas.Add key, 0#
        areas(key) = CDbl(areas(key)) + det / 2#
    Next e
    Set FMAreas = areas
End Function
Private Function FMAliased(ByRef source() As Long, ByRef target() As Long) As Boolean
    On Error Resume Next
    FMAliased = (VarPtr(source(0, 0)) = VarPtr(target(0, 0)))
    err.Clear: On Error GoTo 0
End Function

Public Function RPX_FlowMeshCandidate(ByRef points() As Double, ByRef cells() As Long, ByRef materialId() As Long, ByRef physicalVelocity() As Double, ByVal protectedEdges As Object, ByRef candidateTriangles() As Long, ByRef info As Object, Optional ByVal passes As Long = 5) As Boolean
    Dim nn As Long, ne As Long, i As Long, e As Long, f As Long, k As Long, sweep As Long, completed As Long
    Dim current() As Long, metrics() As FMMetric, identity As FMMetric, average As FMMetric
    Dim edges() As FMEdge, edgeCount As Long, lookup As Object, guards As Object, originalProtected As Object
    Dim choices() As FMChoice, choiceCount As Long, order() As Long, used() As Boolean, pick As Long
    Dim a As Long, b As Long, c As Long, d As Long, p0 As Long, p1 As Long, p2 As Long, q0 As Long, q1 As Long, q2 As Long
    Dim before As Double, after As Double, geometricBefore As Double, geometricAfter As Double, oldArea As Double, newArea As Double
    Dim key As Variant, text As String, pair As Variant, areasBefore As Object, areasAfter As Object, change As Object, changes As Collection
    If passes < 0 Or passes > 5 Then err.Raise 5, "RPX_FlowMesh", "FLOW_PASSES_MUST_BE_0_TO_5"
    If protectedEdges Is Nothing Then err.Raise 5, "RPX_FlowMesh", "FLOW_EXPLICIT_PROTECTED_EDGE_SET_REQUIRED"
    If LBound(points, 1) <> 0 Or LBound(points, 2) <> 0 Or UBound(points, 2) <> 1 Then err.Raise 5, "RPX_FlowMesh", "FLOW_POINT_LAYOUT"
    If LBound(cells, 1) <> 0 Or LBound(cells, 2) <> 0 Or UBound(cells, 2) <> 2 Then err.Raise 5, "RPX_FlowMesh", "FLOW_TRIANGLE_LAYOUT"
    nn = UBound(points, 1) + 1: ne = UBound(cells, 1) + 1
    If nn < 3 Or ne < 1 Or LBound(materialId) <> 0 Or UBound(materialId) <> ne - 1 Then err.Raise 5, "RPX_FlowMesh", "FLOW_MATERIAL_LAYOUT"
    If LBound(physicalVelocity) <> 0 Or UBound(physicalVelocity) <> 12 * ne - 1 Then err.Raise 5, "RPX_FlowMesh", "FLOW_EXACT_PHYSICAL_VELOCITY_LAYOUT"
    If FMAliased(cells, candidateTriangles) Then err.Raise 5, "RPX_FlowMesh", "FLOW_OUTPUT_MUST_NOT_ALIAS_SOURCE"
    For i = 0 To nn - 1
        If Not FMFinite(points(i, 0)) Or Not FMFinite(points(i, 1)) Then err.Raise 5, "RPX_FlowMesh", "FLOW_NONFINITE_POINT"
    Next i
    For e = 0 To ne - 1
        If materialId(e) < 0 Then err.Raise 5, "RPX_FlowMesh", "FLOW_NEGATIVE_MATERIAL_ID"
        For k = 0 To 2
            If cells(e, k) < 0 Or cells(e, k) >= nn Then err.Raise 5, "RPX_FlowMesh", "FLOW_NODE_ID_OUT_OF_RANGE"
        Next k
    Next e
    For i = 0 To 12 * ne - 1
        If Not FMFinite(physicalVelocity(i)) Then err.Raise 5, "RPX_FlowMesh", "FLOW_NONFINITE_VELOCITY"
    Next i
    current = cells: Set areasBefore = FMAreas(points, current, materialId)
    Set lookup = FMEdges(current, edges, edgeCount): Set guards = CreateObject("Scripting.Dictionary")
    Set originalProtected = CreateObject("Scripting.Dictionary")
    For i = 0 To edgeCount - 1
        e = edges(i).e: f = edges(i).f: text = FMKey(edges(i).a, edges(i).b)
        If f < 0 Then
            guards.Add text, True: originalProtected.Add text, "exterior"
        ElseIf materialId(e) <> materialId(f) Then
            guards.Add text, True: originalProtected.Add text, "material"
        End If
    Next i
    For Each key In protectedEdges.Keys
        text = CStr(key): pair = Split(text, ":")
        If UBound(pair) <> 1 Then err.Raise 5, "RPX_FlowMesh", "FLOW_PROTECTED_EDGE_KEY"
        a = CLng(pair(0)): b = CLng(pair(1))
        If a < 0 Or b >= nn Or a >= b Or text <> FMKey(a, b) Or Not lookup.Exists(text) Then err.Raise 5, "RPX_FlowMesh", "FLOW_PROTECTED_EDGE_NOT_IN_SOURCE"
        guards(text) = True
    Next key
    Call FMSourceMetrics(points, cells, physicalVelocity, metrics)
    identity.xx = 1#: identity.xy = 0#: identity.yy = 1#: Set changes = New Collection
    For sweep = 1 To passes
        Set lookup = FMEdges(current, edges, edgeCount)
        ReDim choices(0 To edgeCount - 1): choiceCount = 0
        For i = 0 To edgeCount - 1
            e = edges(i).e: f = edges(i).f: a = edges(i).a: b = edges(i).b
            If f < 0 Then GoTo NextEdge
            If materialId(e) <> materialId(f) Or guards.Exists(FMKey(a, b)) Then GoTo NextEdge
            c = FMOpposite(current, e, a, b): d = FMOpposite(current, f, a, b)
            If c = d Or lookup.Exists(FMKey(c, d)) Then GoTo NextEdge
            If FMDet(points, a, b, c) * FMDet(points, a, b, d) >= 0# Then GoTo NextEdge
            If FMDet(points, c, d, a) * FMDet(points, c, d, b) >= 0# Then GoTo NextEdge
            FMOrient points, c, d, a, p0, p1, p2: FMOrient points, d, c, b, q0, q1, q2
            oldArea = FMDet(points, current(e, 0), current(e, 1), current(e, 2)) + FMDet(points, current(f, 0), current(f, 1), current(f, 2))
            newArea = FMDet(points, p0, p1, p2) + FMDet(points, q0, q1, q2)
            If Abs(oldArea - newArea) > 0.0000000001 Then GoTo NextEdge
            average.xx = (metrics(e).xx + metrics(f).xx) / 2#: average.xy = (metrics(e).xy + metrics(f).xy) / 2#: average.yy = (metrics(e).yy + metrics(f).yy) / 2#
            before = FMMin(FMQuality(points, current(e, 0), current(e, 1), current(e, 2), average), FMQuality(points, current(f, 0), current(f, 1), current(f, 2), average))
            after = FMMin(FMQuality(points, p0, p1, p2, average), FMQuality(points, q0, q1, q2, average))
            geometricBefore = FMMin(FMQuality(points, current(e, 0), current(e, 1), current(e, 2), identity), FMQuality(points, current(f, 0), current(f, 1), current(f, 2), identity))
            geometricAfter = FMMin(FMQuality(points, p0, p1, p2, identity), FMQuality(points, q0, q1, q2, identity))
            If after > 1.05 * before + 0.0000000001 And geometricAfter >= 0.6 * geometricBefore Then
                With choices(choiceCount)
                    .e = e: .f = f: .a = a: .b = b: .c = c: .d = d: .p0 = p0: .p1 = p1: .p2 = p2: .q0 = q0: .q1 = q1: .q2 = q2
                    .gain = after - before: .ordinal = i
                End With
                choiceCount = choiceCount + 1
            End If
NextEdge:
        Next i
        If choiceCount = 0 Then Exit For
        ReDim order(0 To choiceCount - 1): ReDim used(0 To ne - 1)
        For i = 0 To choiceCount - 1: order(i) = i: Next i
        FMSort order, choices, 0, choiceCount - 1
        For i = 0 To choiceCount - 1
            pick = order(i): e = choices(pick).e: f = choices(pick).f
            If Not used(e) And Not used(f) Then
                With choices(pick)
                    current(e, 0) = .p0: current(e, 1) = .p1: current(e, 2) = .p2
                    current(f, 0) = .q0: current(f, 1) = .q1: current(f, 2) = .q2
                    Set change = CreateObject("Scripting.Dictionary")
                    change.Add "old", Array(.a, .b): change.Add "new", Array(.c, .d): change.Add "gain", .gain
                    change.Add "sweep", sweep: change.Add "elements", Array(e, f): changes.Add change
                End With
                used(e) = True: used(f) = True
            End If
        Next i
        completed = sweep
    Next sweep
    Set lookup = FMEdges(current, edges, edgeCount)
    For Each key In guards.Keys
        If Not lookup.Exists(CStr(key)) Then err.Raise 5, "RPX_FlowMesh", "FLOW_PROTECTED_EDGE_CHANGED"
    Next key
    For i = 0 To edgeCount - 1
        e = edges(i).e: f = edges(i).f: text = FMKey(edges(i).a, edges(i).b)
        If f < 0 Then
            If Not originalProtected.Exists(text) Then err.Raise 5, "RPX_FlowMesh", "FLOW_EXTERIOR_EDGE_CHANGED"
            If CStr(originalProtected(text)) <> "exterior" Then err.Raise 5, "RPX_FlowMesh", "FLOW_EXTERIOR_EDGE_CLASS_CHANGED"
        ElseIf materialId(e) <> materialId(f) Then
            If Not originalProtected.Exists(text) Then err.Raise 5, "RPX_FlowMesh", "FLOW_MATERIAL_EDGE_CHANGED"
            If CStr(originalProtected(text)) <> "material" Then err.Raise 5, "RPX_FlowMesh", "FLOW_MATERIAL_EDGE_CLASS_CHANGED"
        End If
    Next i
    Set areasAfter = FMAreas(points, current, materialId)
    For Each key In areasBefore.Keys
        If Abs(CDbl(areasBefore(key)) - CDbl(areasAfter(key))) > 0.0000000001 + 0.000000000001 * Abs(CDbl(areasBefore(key))) Then err.Raise 5, "RPX_FlowMesh", "FLOW_MATERIAL_AREA_CHANGED"
    Next key
    candidateTriangles = current: Set info = CreateObject("Scripting.Dictionary")
    info.Add "method", "flow_metric_flip": info.Add "certificate", False: info.Add "nested", False
    info.Add "passes", passes: info.Add "completed_passes", completed: info.Add "flips", changes.count
    info.Add "elements_before", ne: info.Add "elements_after", ne: info.Add "protected_edges", guards.count
    info.Add "changes", changes: info.Add "material_areas_before", areasBefore: info.Add "material_areas_after", areasAfter
    info.Add "source_metrics_fixed_by_element_id", True: info.Add "requires_new_original_q8_field_and_root", True
    RPX_FlowMeshCandidate = True
End Function
