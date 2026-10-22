Attribute VB_Name = "RPX_Refine"
Option Explicit
Private locXY() As Double, locTri() As Long, locNE As Long
Private locX0 As Double, locY0 As Double, locW As Double, locH As Double, locNX As Long, locNY As Long
Private locStart() As Long, locList() As Long
' Conforming subdivision, preserving every old triangle and material interface.
' The old P2 fields restrict to P2 fields on the subtriangles.
Public Sub RPX_RefineMesh(ByRef indicator() As Double, ByVal fraction As Double, Optional ByVal forceLegacy As Boolean = False, Optional ByVal forceRed As Boolean = False)
    Dim marks As Object, key As Variant, rec As Variant, rank() As Long, i As Long, j As Long
    Dim count As Long, e As Long, k As Long, a As Long, b As Long, center As Long, oldRE As Long
    Dim oldTri() As Long, poly(0 To 5) As Long, np As Long, hit As Boolean, nodes As Long
    Dim red As Boolean, mids(0 To 2) As Long, marked As Long, a0 As Long, b0 As Long, c0 As Long, ab As Long, ca As Long
    Dim oldMaterial() As Long, oldRigid() As Long, oldBody0() As Double, oldBody1() As Double
    If forceRed Then
        If RR <> 0 Or RPX_WaterCouplingOn() Or fraction <> 1# Then err.Raise 5, "RPX_Refine", "UNIFORM_RED_SCOPE_UNSUPPORTED"
        oldMaterial = RMatId: oldRigid = RRigidId: oldBody0 = RBody0: oldBody1 = RBody1
    End If
    red = (RPX_RobustMeshPolicy() Or forceRed) And Not forceLegacy
    If fraction <= 0# Or fraction > 1# Then err.Raise 5, , "REFINE_FRACTION_OUT_OF_RANGE"
    RProgramKind = "": oldRE = re: oldTri = RTri
    RefineRank indicator, rank
    count = CLng(re * fraction): If count < 1 Then count = 1
    Set marks = RPX_Row()
    For i = 0 To count - 1
        e = rank(i)
        For k = 0 To 2
            a = RTri(e, k): b = RTri(e, (k + 1) Mod 3)
            key = RPX_EdgeKey(a, b)
            If Not marks.Exists(key) Then
                If rn >= UBound(RXY, 1) - 3 Then err.Raise 7, , "REFINE_NODE_CAPACITY"
                marks(key) = rn
                RXY(rn, 0) = (RXY(a, 0) + RXY(b, 0)) / 2#: RXY(rn, 1) = (RXY(a, 1) + RXY(b, 1)) / 2#
                If RConstraints.Exists(key) Then
                    RConstraints.Remove key
                    RConstraints(RPX_EdgeKey(a, rn)) = Array(a, rn)
                    RConstraints(RPX_EdgeKey(rn, b)) = Array(rn, b)
                End If
                rn = rn + 1
            End If
        Next k
    Next i
    re = 0
    For e = 0 To oldRE - 1
        np = 0: hit = False: marked = 0
        For k = 0 To 2: mids(k) = -1: Next k
        For k = 0 To 2
            a = oldTri(e, k): b = oldTri(e, (k + 1) Mod 3)
            poly(np) = a: np = np + 1: key = RPX_EdgeKey(a, b)
            If marks.Exists(key) Then poly(np) = marks(key): mids(k) = marks(key): np = np + 1: hit = True: marked = marked + 1
        Next k
        If Not hit Then
            For k = 0 To 2: RTri(re, k) = oldTri(e, k): Next k
            re = re + 1
        ElseIf red Then
            If marked = 3 Then
                RefineChild oldTri(e, 0), mids(0), mids(2)
                RefineChild mids(0), oldTri(e, 1), mids(1)
                RefineChild mids(2), mids(1), oldTri(e, 2)
                RefineChild mids(0), mids(1), mids(2)
            ElseIf marked = 1 Then
                For k = 0 To 2: If mids(k) >= 0 Then Exit For
                Next k
                a0 = oldTri(e, k): b0 = oldTri(e, (k + 1) Mod 3): c0 = oldTri(e, (k + 2) Mod 3)
                RefineChild a0, mids(k), c0
                RefineChild mids(k), b0, c0
            Else
                For k = 0 To 2: If mids(k) >= 0 And mids((k + 2) Mod 3) >= 0 Then Exit For
                Next k
                a0 = oldTri(e, k): b0 = oldTri(e, (k + 1) Mod 3): c0 = oldTri(e, (k + 2) Mod 3)
                ab = mids(k): ca = mids((k + 2) Mod 3)
                RefineChild a0, ab, ca
                If (RXY(ab, 0) - RXY(c0, 0)) ^ 2 + (RXY(ab, 1) - RXY(c0, 1)) ^ 2 <= (RXY(ca, 0) - RXY(b0, 0)) ^ 2 + (RXY(ca, 1) - RXY(b0, 1)) ^ 2 Then
                    RefineChild ab, b0, c0: RefineChild ab, c0, ca
                Else
                    RefineChild ab, b0, ca: RefineChild b0, c0, ca
                End If
            End If
        Else
            center = rn
            If rn >= UBound(RXY, 1) - 3 Then err.Raise 7, , "REFINE_NODE_CAPACITY"
            RXY(rn, 0) = 0#: RXY(rn, 1) = 0#
            For k = 0 To 2
                RXY(rn, 0) = RXY(rn, 0) + RXY(oldTri(e, k), 0) / 3#
                RXY(rn, 1) = RXY(rn, 1) + RXY(oldTri(e, k), 1) / 3#
            Next k
            rn = rn + 1
            For k = 0 To np - 1
                If re > UBound(RTri, 1) Then err.Raise 7, , "REFINE_ELEMENT_CAPACITY"
                RPX_SetPositive re, poly(k), poly((k + 1) Mod np), center: re = re + 1
            Next k
        End If
    Next e
    If forceRed Then
        If re <> 4 * oldRE Then err.Raise 5, "RPX_Refine", "UNIFORM_RED_CHILD_COUNT"
        ReDim RMatId(0 To re - 1): ReDim RRigidId(0 To re - 1)
        ReDim RBody0(0 To re - 1, 0 To 1): ReDim RBody1(0 To re - 1, 0 To 1)
        For e = 0 To oldRE - 1: For k = 0 To 3
            i = 4 * e + k: RMatId(i) = oldMaterial(e): RRigidId(i) = oldRigid(e)
            For j = 0 To 1: RBody0(i, j) = oldBody0(e, j): RBody1(i, j) = oldBody1(e, j): Next j
        Next k: Next e
        ' Exact subdivision preserves parent material/body. Do not insert the
        ' independent free-corner/fan repair geometry after the red4 operation.
        Call RPX_Topology: Call RPX_MapFaces
    Else
        Call RPX_AssignRegions
    End If
End Sub
Public Function RPX_TestRefine() As String
    Dim eta() As Double, e As Long, area As Double, oldElements As Long
    RPX_SlopeDefinition 64: Call RPX_GenerateMesh
    oldElements = re: ReDim eta(0 To re - 1)
    For e = 0 To re - 1: eta(e) = e + 1#: Next e
    RPX_RefineMesh eta, 0.2
    For e = 0 To re - 1: area = area + RArea(e): Next e
    If Abs(area - 1650#) > 0.000001 Or re <= oldElements Then err.Raise 5, , "REFINE_AREA_FAILED"
    For e = 0 To RNEdge - 1
        If REdges(e).element(1) < 0 Then
            If Not RConstraints.Exists(RPX_EdgeKey(REdges(e).a, REdges(e).b)) Then err.Raise 5, , "REFINE_HANGING_EDGE"
        End If
    Next e
    RPX_TestRefine = "PASS area=" & CStr(area) & " elements=" & CStr(oldElements) & " -> " & CStr(re)
End Function

Public Function RPX_RobustMeshPolicy() As Boolean
    Dim policy As String
    policy = UCase$(Trim$(CStr(RPX_Setting("ADAPT_MESH_POLICY", "LEGACY"))))
    Select Case policy
        Case "LEGACY"
        Case "ROBUST_REFINE", "LEGACY_THEN_REFINE": RPX_RobustMeshPolicy = (RR = 0)
        Case Else: err.Raise 5, "RPX_Refine", "INVALID_ADAPT_MESH_POLICY"
    End Select
End Function
Private Sub RefineChild(ByVal a As Long, ByVal b As Long, ByVal c As Long)
    If re > UBound(RTri, 1) Then err.Raise 7, , "REFINE_ELEMENT_CAPACITY"
    RPX_SetPositive re, a, b, c: re = re + 1
End Sub
Private Sub RefineRank(ByRef indicator() As Double, ByRef rank() As Long)
    Dim e As Long, i As Long, j As Long
    ReDim rank(0 To re - 1)
    For e = 0 To re - 1: rank(e) = e: Next e
    ' Stable insertion sort is sufficient for the initial 1536-element density.
    For i = 1 To re - 1
        e = rank(i): j = i - 1
        Do While j >= 0
            If indicator(rank(j)) >= indicator(e) Then Exit Do
            rank(j + 1) = rank(j): j = j - 1
        Loop
        rank(j + 1) = e
    Next i
End Sub
Public Function RPX_RefineBudgetFraction(ByRef indicator() As Double, ByVal cap As Long, ByVal maximum As Double) As Double
    Dim rank() As Long, costs As Object, marks As Object, key As String, e As Long, k As Long, i As Long
    Dim count As Long, selected As Long, needed As Long, elements As Long, a As Long, b As Long
    If cap <= re Then Exit Function
    RefineRank indicator, rank
    Set costs = RPX_Row(): Set marks = RPX_Row()
    For i = 0 To RNEdge - 1
        key = RPX_EdgeKey(REdges(i).a, REdges(i).b)
        If REdges(i).element(1) < 0 Then costs(key) = 1 Else costs(key) = 2
    Next i
    count = CLng(re * maximum): If count < 1 Then count = 1
    elements = re
    For i = 0 To count - 1
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        e = rank(i): needed = 0
        For k = 0 To 2
            a = RTri(e, k): b = RTri(e, (k + 1) Mod 3): key = RPX_EdgeKey(a, b)
            If Not marks.Exists(key) Then needed = needed + CLng(costs(key))
        Next k
        If elements + needed > cap Then Exit For
        For k = 0 To 2: key = RPX_EdgeKey(RTri(e, k), RTri(e, (k + 1) Mod 3)): marks(key) = 1: Next k
        elements = elements + needed: selected = selected + 1
    Next i
    If selected > 0 Then
        RPX_RefineBudgetFraction = selected / CDbl(re)
        If CLng(re * RPX_RefineBudgetFraction) <> selected Then err.Raise 5, , "REFINE_BUDGET_COUNT_MISMATCH"
    End If
    RPX_DiagEvent "adapt_refine_budget", "cap=" & cap & ";predicted_elements=" & elements & ";selected=" & selected
End Function
' ---------------------------------------------------------------------------
' G1R20 Fix8 (candidate): thin-layer mesh operators ported from the P06 Python
' prototype (validation/G1R20_P06_meshproto_20261018). Geometry only: no field
' crosses a mesh as a witness and every new bound still needs its own original
' q8 field, full audit and own root. Rigid and water models are not supported.
' ---------------------------------------------------------------------------

' thick(m) = 2*A/(P/2)^2 with A the material area and P the length of its
' exterior and material-interface edges (mesh independent geometry measure).
Public Sub RPX_MaterialThickness(ByRef thick() As Double)
    Dim area() As Double, perimeter() As Double, id As Long, e As Long, f As Long, m As Long, half As Double
    ReDim area(0 To RM - 1): ReDim perimeter(0 To RM - 1): ReDim thick(0 To RM - 1)
    For e = 0 To re - 1: area(RMatId(e)) = area(RMatId(e)) + RArea(e): Next e
    For id = 0 To RNEdge - 1
        e = REdges(id).element(0): f = REdges(id).element(1)
        If f < 0 Then
            perimeter(RMatId(e)) = perimeter(RMatId(e)) + REdges(id).length
        ElseIf RMatId(e) <> RMatId(f) Then
            perimeter(RMatId(e)) = perimeter(RMatId(e)) + REdges(id).length
            perimeter(RMatId(f)) = perimeter(RMatId(f)) + REdges(id).length
        End If
    Next id
    For m = 0 To RM - 1
        half = perimeter(m) / 2#
        If area(m) > 0# And half > 0# Then thick(m) = 2# * area(m) / (half * half) Else thick(m) = 1E+30
    Next m
End Sub

' Interior same-material edges (material in mask) whose two end nodes both lie
' on that material's boundary. In a one-element-thick layer they span the layer;
' splitting them creates a continuous same-material mid-line.
Public Function RPX_TransverseEdges(ByRef mask() As Boolean, ByVal marks As Object) As Long
    Dim onb As Object, id As Long, e As Long, f As Long, m As Long, key As String
    Set onb = RPX_Row()
    For id = 0 To RNEdge - 1
        e = REdges(id).element(0): f = REdges(id).element(1)
        If f < 0 Then
            onb(REdges(id).a & "|" & RMatId(e)) = 1: onb(REdges(id).b & "|" & RMatId(e)) = 1
        ElseIf RMatId(e) <> RMatId(f) Then
            onb(REdges(id).a & "|" & RMatId(e)) = 1: onb(REdges(id).b & "|" & RMatId(e)) = 1
            onb(REdges(id).a & "|" & RMatId(f)) = 1: onb(REdges(id).b & "|" & RMatId(f)) = 1
        End If
    Next id
    For id = 0 To RNEdge - 1
        e = REdges(id).element(0): f = REdges(id).element(1)
        If f >= 0 Then
            m = RMatId(e)
            If m = RMatId(f) Then
                If mask(m) Then
                    If onb.Exists(REdges(id).a & "|" & m) And onb.Exists(REdges(id).b & "|" & m) Then
                        key = RPX_EdgeKey(REdges(id).a, REdges(id).b)
                        If Not marks.Exists(key) Then marks(key) = 1: RPX_TransverseEdges = RPX_TransverseEdges + 1
                    End If
                End If
            End If
        End If
    Next id
End Function

' Upper kinematic score (Python experiment_core.indicator): vertex shear-strain
' magnitude plus Simpson velocity jumps on every interior edge.
Public Sub RPX_ThinUpperScore(ByRef velocity() As Double, ByRef w() As Double)
    Dim e As Long, f As Long, k As Long, j As Long, id As Long, q As Long
    Dim bary(0 To 2) As Double, g(0 To 5, 0 To 1) As Double, tw(0 To 2) As Double
    Dim ux As Double, uy As Double, vx As Double, vy As Double, s As Double, jx As Double, jy As Double, value As Double
    ReDim w(0 To re - 1)
    If UBound(velocity) < 12 * re - 1 Then err.Raise 5, "RPX_Refine", "FIX8_UPPER_SCORE_LAYOUT"
    For e = 0 To re - 1
        For k = 0 To 2
            For j = 0 To 2: bary(j) = 0#: Next j
            bary(k) = 1#
            RPX_P2Grad e, bary, g
            ux = 0#: uy = 0#: vx = 0#: vy = 0#
            For j = 0 To 5
                ux = ux + velocity(12 * e + 2 * j) * g(j, 0): uy = uy + velocity(12 * e + 2 * j) * g(j, 1)
                vx = vx + velocity(12 * e + 2 * j + 1) * g(j, 0): vy = vy + velocity(12 * e + 2 * j + 1) * g(j, 1)
            Next j
            w(e) = w(e) + RArea(e) * Sqr((ux - vy) ^ 2 + (uy + vx) ^ 2) / 3#
        Next k
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1)
            If f >= 0 Then
                value = 0#
                For q = 0 To 2
                    s = 0.5 * q: RPX_TraceWeights s, tw
                    jx = 0#: jy = 0#
                    For k = 0 To 2
                        jx = jx + tw(k) * (velocity(12 * f + 2 * .loc(1, k)) - velocity(12 * e + 2 * .loc(0, k)))
                        jy = jy + tw(k) * (velocity(12 * f + 2 * .loc(1, k) + 1) - velocity(12 * e + 2 * .loc(0, k) + 1))
                    Next k
                    If q = 1 Then value = value + 4# * Sqr(jx * jx + jy * jy) Else value = value + Sqr(jx * jx + jy * jy)
                Next q
                value = .length * value / 6#
                w(e) = w(e) + value / 2#: w(f) = w(f) + value / 2#
            End If
        End With
    Next id
End Sub

' Lower score (Python mesh_accuracy.lower_stress_score, proximity background 0):
' Bernstein stress variation x yield proximity plus same-material stress jumps.
Public Sub RPX_ThinLowerScore(ByRef field() As Double, ByRef w() As Double)
    Dim e As Long, f As Long, k As Long, j As Long, id As Long, count As Long
    Dim reference As Double, total As Double, mean(0 To 2) As Double, sg(0 To 5, 0 To 2) As Double
    Dim c As Double, phi As Double, sinPhi As Double, cosPhi As Double, tanPhi As Double, cap As Double, margin As Double
    Dim variation As Double, proximity() As Double, jump As Double, term As Double, d As Double
    ReDim w(0 To re - 1): ReDim proximity(0 To re - 1)
    For e = 0 To re - 1
        If RRigidId(e) >= 0 Then err.Raise 5, "RPX_Refine", "FIX8_LOWER_SCORE_RIGID"
        For k = 0 To 5: For j = 0 To 2
            d = field(RPX_StressId(e, k, j)): total = total + d * d
        Next j: Next k
        count = count + 6
    Next e
    reference = 0.000000000001: If count > 0 Then If Sqr(total / count) > reference Then reference = Sqr(total / count)
    For e = 0 To re - 1
        For j = 0 To 2: mean(j) = 0#: Next j
        For k = 0 To 5: For j = 0 To 2
            sg(k, j) = field(RPX_StressId(e, k, j)): mean(j) = mean(j) + sg(k, j) / 6#
        Next j: Next k
        variation = 0#
        For k = 0 To 5: For j = 0 To 2: variation = variation + (sg(k, j) - mean(j)) ^ 2: Next j: Next k
        variation = variation / 6# / (reference * reference)
        RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
        cap = 2# * c * cosPhi
        For k = 0 To 5
            margin = cap - Sqr((sg(k, 0) - sg(k, 1)) ^ 2 + 4# * sg(k, 2) ^ 2) - sinPhi * (sg(k, 0) + sg(k, 1))
            If margin < 0# Then margin = 0#
            proximity(e) = proximity(e) + Exp(-8# * margin / (reference + cap)) / 6#
        Next k
        w(e) = RArea(e) * variation * proximity(e)
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1)
            If f >= 0 Then
                If RMatId(e) = RMatId(f) Then
                    jump = 0#
                    For k = 0 To 2: For j = 0 To 2
                        jump = jump + (field(RPX_StressId(e, .loc(0, k), j)) - field(RPX_StressId(f, .loc(1, k), j))) ^ 2
                    Next j: Next k
                    jump = jump / 3# / (reference * reference)
                    term = 0.25 * (RArea(e) + RArea(f)) / 2# * jump * (proximity(e) + proximity(f)) / 2#
                    w(e) = w(e) + term / 2#: w(f) = w(f) + term / 2#
                End If
            End If
        End With
    Next id
End Sub

' Doerfler marking: smallest top set whose score reaches theta of the total.
Public Sub RPX_DoerflerMark(ByRef score() As Double, ByVal theta As Double, ByRef marked() As Boolean)
    Dim order() As Long, e As Long, n As Long, total As Double, cum As Double, i As Long
    n = UBound(score) + 1: ReDim marked(0 To n - 1): ReDim order(0 To n - 1)
    For e = 0 To n - 1
        If Not (score(e) >= 0#) Then err.Raise 5, "RPX_Refine", "FIX8_SCORE_NEGATIVE_OR_NAN"
        order(e) = e: total = total + score(e)
    Next e
    If total <= 0# Then
        For e = 0 To n - 1: marked(e) = True: Next e
        Exit Sub
    End If
    If n > 1 Then ScoreSort order, score, 0, n - 1
    For i = 0 To n - 1
        marked(order(i)) = True: cum = cum + score(order(i))
        If cum >= theta * total Then Exit For
    Next i
End Sub
Private Sub ScoreSort(ByRef order() As Long, ByRef score() As Double, ByVal first As Long, ByVal last As Long)
    Dim left As Long, right As Long, pivot As Double, t As Long
    left = first: right = last: pivot = score(order((first + last) \ 2))
    Do While left <= right
        Do While score(order(left)) > pivot: left = left + 1: Loop
        Do While score(order(right)) < pivot: right = right - 1: Loop
        If left <= right Then
            t = order(left): order(left) = order(right): order(right) = t
            left = left + 1: right = right - 1
        End If
    Loop
    If first < right Then ScoreSort order, score, first, right
    If left < last Then ScoreSort order, score, left, last
End Sub

' One edge per marked element: the edge of largest length in the metric
' J'J + 0.02 tr(J'J) I, J the P2 velocity gradient at the centroid.
Public Function RPX_DirectionalEdges(ByRef velocity() As Double, ByRef score() As Double, ByVal theta As Double, ByVal marks As Object) As Long
    Dim marked() As Boolean, e As Long, j As Long, k As Long, best As Long, key As String
    Dim bary(0 To 2) As Double, g(0 To 5, 0 To 1) As Double
    Dim ux As Double, uy As Double, vx As Double, vy As Double, m11 As Double, m12 As Double, m22 As Double, reg As Double
    Dim dx As Double, dy As Double, value As Double, top As Double
    RPX_DoerflerMark score, theta, marked
    For j = 0 To 2: bary(j) = 1# / 3#: Next j
    For e = 0 To re - 1
        If marked(e) Then
            RPX_P2Grad e, bary, g
            ux = 0#: uy = 0#: vx = 0#: vy = 0#
            For j = 0 To 5
                ux = ux + velocity(12 * e + 2 * j) * g(j, 0): uy = uy + velocity(12 * e + 2 * j) * g(j, 1)
                vx = vx + velocity(12 * e + 2 * j + 1) * g(j, 0): vy = vy + velocity(12 * e + 2 * j + 1) * g(j, 1)
            Next j
            m11 = ux * ux + vx * vx: m12 = ux * uy + vx * vy: m22 = uy * uy + vy * vy
            reg = 0.02 * (m11 + m22): If reg < 1E-18 Then reg = 1E-18
            m11 = m11 + reg: m22 = m22 + reg
            best = 0: top = -1#
            For k = 0 To 2
                dx = RXY(RTri(e, (k + 1) Mod 3), 0) - RXY(RTri(e, k), 0): dy = RXY(RTri(e, (k + 1) Mod 3), 1) - RXY(RTri(e, k), 1)
                value = dx * dx * m11 + 2# * dx * dy * m12 + dy * dy * m22
                If value > top Then top = value: best = k
            Next k
            key = RPX_EdgeKey(RTri(e, best), RTri(e, (best + 1) Mod 3))
            If Not marks.Exists(key) Then marks(key) = 1: RPX_DirectionalEdges = RPX_DirectionalEdges + 1
        End If
    Next e
End Function

' All three edges of each marked element.
Public Function RPX_RedEdges(ByRef marked() As Boolean, ByVal marks As Object) As Long
    Dim e As Long, k As Long, key As String
    For e = 0 To re - 1
        If marked(e) Then
            For k = 0 To 2
                key = RPX_EdgeKey(RTri(e, k), RTri(e, (k + 1) Mod 3))
                If Not marks.Exists(key) Then marks(key) = 1: RPX_RedEdges = RPX_RedEdges + 1
            Next k
        End If
    Next e
End Function

' Split exactly the given edges at their midpoints (red/green 1-, 2-, 3-edge
' patterns, same as RPX_RefineMesh red). Children keep the parent material,
' rigid id and body force; boundary faces are mapped again from the inputs.
' The caller converts loads to the reference phase. RXY/RTri need capacity.
Public Sub RPX_SplitEdgeSet(ByVal edgeKeys As Object)
    Dim mids As Object, key As Variant, pair As Variant, a As Long, b As Long, e As Long, k As Long, j As Long
    Dim oldRE As Long, oldTri() As Long, oldMaterial() As Long, oldRigid() As Long, oldBody0() As Double, oldBody1() As Double
    Dim parent() As Long, md(0 To 2) As Long, marked As Long, a0 As Long, b0 As Long, c0 As Long, ab As Long, ca As Long, first As Long, child As Long
    If RR <> 0 Then err.Raise 5, "RPX_Refine", "FIX8_SPLIT_RIGID_UNSUPPORTED"
    If edgeKeys.count = 0 Then Exit Sub
    oldRE = re: oldTri = RTri: oldMaterial = RMatId: oldRigid = RRigidId: oldBody0 = RBody0: oldBody1 = RBody1
    Set mids = RPX_Row()
    For Each key In edgeKeys.keys
        pair = Split(CStr(key), ":"): a = CLng(pair(0)): b = CLng(pair(1))
        If rn >= UBound(RXY, 1) - 3 Then err.Raise 7, "RPX_Refine", "REFINE_NODE_CAPACITY"
        mids(CStr(key)) = rn
        RXY(rn, 0) = (RXY(a, 0) + RXY(b, 0)) / 2#: RXY(rn, 1) = (RXY(a, 1) + RXY(b, 1)) / 2#
        If RConstraints.Exists(CStr(key)) Then
            RConstraints.Remove CStr(key)
            RConstraints(RPX_EdgeKey(a, rn)) = Array(a, rn)
            RConstraints(RPX_EdgeKey(rn, b)) = Array(rn, b)
        End If
        rn = rn + 1
    Next key
    ReDim parent(0 To UBound(RTri, 1))
    re = 0
    For e = 0 To oldRE - 1
        first = re: marked = 0
        For k = 0 To 2
            md(k) = -1: key = RPX_EdgeKey(oldTri(e, k), oldTri(e, (k + 1) Mod 3))
            If mids.Exists(key) Then md(k) = mids(key): marked = marked + 1
        Next k
        If marked = 0 Then
            RefineChild oldTri(e, 0), oldTri(e, 1), oldTri(e, 2)
        ElseIf marked = 3 Then
            RefineChild oldTri(e, 0), md(0), md(2)
            RefineChild md(0), oldTri(e, 1), md(1)
            RefineChild md(2), md(1), oldTri(e, 2)
            RefineChild md(0), md(1), md(2)
        ElseIf marked = 1 Then
            For k = 0 To 2: If md(k) >= 0 Then Exit For
            Next k
            a0 = oldTri(e, k): b0 = oldTri(e, (k + 1) Mod 3): c0 = oldTri(e, (k + 2) Mod 3)
            RefineChild a0, md(k), c0
            RefineChild md(k), b0, c0
        Else
            For k = 0 To 2: If md(k) >= 0 And md((k + 2) Mod 3) >= 0 Then Exit For
            Next k
            a0 = oldTri(e, k): b0 = oldTri(e, (k + 1) Mod 3): c0 = oldTri(e, (k + 2) Mod 3)
            ab = md(k): ca = md((k + 2) Mod 3)
            RefineChild a0, ab, ca
            If (RXY(ab, 0) - RXY(c0, 0)) ^ 2 + (RXY(ab, 1) - RXY(c0, 1)) ^ 2 <= (RXY(ca, 0) - RXY(b0, 0)) ^ 2 + (RXY(ca, 1) - RXY(b0, 1)) ^ 2 Then
                RefineChild ab, b0, c0: RefineChild ab, c0, ca
            Else
                RefineChild ab, b0, ca: RefineChild b0, c0, ca
            End If
        End If
        For child = first To re - 1: parent(child) = e: Next child
    Next e
    ReDim RMatId(0 To re - 1): ReDim RRigidId(0 To re - 1)
    ReDim RBody0(0 To re - 1, 0 To 1): ReDim RBody1(0 To re - 1, 0 To 1)
    For child = 0 To re - 1
        e = parent(child): RMatId(child) = oldMaterial(e): RRigidId(child) = oldRigid(e)
        For j = 0 To 1: RBody0(child, j) = oldBody0(e, j): RBody1(child, j) = oldBody1(e, j): Next j
    Next child
    RProgramKind = ""
    Call RPX_Topology: Call RPX_MapFaces
End Sub

' Point location on a saved triangle mesh (uniform bucket grid).
Public Sub RPX_LocateBuild(ByRef xy() As Double, ByRef tri() As Long, ByVal ne As Long)
    Dim e As Long, k As Long, i As Long, j As Long, i0 As Long, i1 As Long, j0 As Long, j1 As Long, pass As Long
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double, xmin As Double, xmax As Double, ymin As Double, ymax As Double
    Dim fill() As Long, cell As Long
    locXY = xy: locTri = tri: locNE = ne
    xmin = 1E+300: ymin = 1E+300: xmax = -1E+300: ymax = -1E+300
    For e = 0 To ne - 1
        For k = 0 To 2
            If xy(tri(e, k), 0) < xmin Then xmin = xy(tri(e, k), 0)
            If xy(tri(e, k), 0) > xmax Then xmax = xy(tri(e, k), 0)
            If xy(tri(e, k), 1) < ymin Then ymin = xy(tri(e, k), 1)
            If xy(tri(e, k), 1) > ymax Then ymax = xy(tri(e, k), 1)
        Next k
    Next e
    locNX = CLng(Sqr(ne)) + 1: locNY = locNX
    locX0 = xmin: locY0 = ymin
    locW = (xmax - xmin) / locNX: If locW <= 0# Then locW = 1#
    locH = (ymax - ymin) / locNY: If locH <= 0# Then locH = 1#
    ReDim locStart(0 To locNX * locNY): ReDim fill(0 To locNX * locNY)
    For pass = 0 To 1
        For e = 0 To ne - 1
            x0 = 1E+300: y0 = 1E+300: x1 = -1E+300: y1 = -1E+300
            For k = 0 To 2
                If xy(tri(e, k), 0) < x0 Then x0 = xy(tri(e, k), 0)
                If xy(tri(e, k), 0) > x1 Then x1 = xy(tri(e, k), 0)
                If xy(tri(e, k), 1) < y0 Then y0 = xy(tri(e, k), 1)
                If xy(tri(e, k), 1) > y1 Then y1 = xy(tri(e, k), 1)
            Next k
            i0 = LocCell(x0, locX0, locW, locNX): i1 = LocCell(x1, locX0, locW, locNX)
            j0 = LocCell(y0, locY0, locH, locNY): j1 = LocCell(y1, locY0, locH, locNY)
            For j = j0 To j1: For i = i0 To i1
                cell = j * locNX + i
                If pass = 0 Then
                    locStart(cell + 1) = locStart(cell + 1) + 1
                Else
                    locList(locStart(cell) + fill(cell)) = e: fill(cell) = fill(cell) + 1
                End If
            Next i: Next j
        Next e
        If pass = 0 Then
            For cell = 1 To locNX * locNY: locStart(cell) = locStart(cell) + locStart(cell - 1): Next cell
            ReDim locList(0 To locStart(locNX * locNY))
        End If
    Next pass
End Sub
Private Function LocCell(ByVal v As Double, ByVal origin As Double, ByVal width As Double, ByVal n As Long) As Long
    Dim i As Long
    i = Int((v - origin) / width)
    If i < 0 Then i = 0
    If i > n - 1 Then i = n - 1
    LocCell = i
End Function
' Returns the saved triangle containing (x, y) and its barycentric coordinates
' (vertex order of that triangle), or -1 when no triangle contains the point.
Public Function RPX_LocatePoint(ByVal x As Double, ByVal y As Double, ByRef lam() As Double) As Long
    Dim cell As Long, p As Long, e As Long, a As Long, b As Long, c As Long, det As Double
    Dim l0 As Double, l1 As Double, l2 As Double, worst As Double, bestWorst As Double, best As Long
    RPX_LocatePoint = -1: best = -1: bestWorst = -1E+300
    cell = LocCell(y, locY0, locH, locNY) * locNX + LocCell(x, locX0, locW, locNX)
    For p = locStart(cell) To locStart(cell + 1) - 1
        e = locList(p): a = locTri(e, 0): b = locTri(e, 1): c = locTri(e, 2)
        det = (locXY(b, 0) - locXY(a, 0)) * (locXY(c, 1) - locXY(a, 1)) - (locXY(c, 0) - locXY(a, 0)) * (locXY(b, 1) - locXY(a, 1))
        If det <> 0# Then
            l0 = ((locXY(b, 0) - x) * (locXY(c, 1) - y) - (locXY(c, 0) - x) * (locXY(b, 1) - y)) / det
            l1 = ((locXY(c, 0) - x) * (locXY(a, 1) - y) - (locXY(a, 0) - x) * (locXY(c, 1) - y)) / det
            l2 = 1# - l0 - l1
            worst = l0: If l1 < worst Then worst = l1
            If l2 < worst Then worst = l2
            If worst > bestWorst Then bestWorst = worst: best = e: lam(0) = l0: lam(1) = l1: lam(2) = l2
        End If
    Next p
    If best >= 0 And bestWorst >= -0.000000001 Then RPX_LocatePoint = best
End Function

' Restriction of a discontinuous P2 velocity on a parent mesh to the current
' (nested child) mesh: exact for nested children. Used only as the flow-flip
' metric source; it is never saved, audited or used as a bound.
Public Function RPX_ProlongUpperVelocity(ByRef pxy() As Double, ByRef ptri() As Long, ByVal pne As Long, ByRef pvel() As Double, ByRef child() As Double) As Boolean
    Dim e As Long, p As Long, k As Long, m As Long, c As Long, cx As Double, cy As Double, x As Double, y As Double
    Dim lam(0 To 2) As Double, node(0 To 5, 0 To 1) As Double, nb(0 To 5) As Double, l(0 To 2) As Double, det As Double
    Dim a As Long, b As Long, cc As Long
    RPX_LocateBuild pxy, ptri, pne
    ReDim child(0 To 12 * re - 1)
    For e = 0 To re - 1
        cx = (RXY(RTri(e, 0), 0) + RXY(RTri(e, 1), 0) + RXY(RTri(e, 2), 0)) / 3#
        cy = (RXY(RTri(e, 0), 1) + RXY(RTri(e, 1), 1) + RXY(RTri(e, 2), 1)) / 3#
        p = RPX_LocatePoint(cx, cy, lam)
        If p < 0 Then Exit Function
        For k = 0 To 2
            node(k, 0) = RXY(RTri(e, k), 0): node(k, 1) = RXY(RTri(e, k), 1)
        Next k
        For k = 0 To 2
            node(3 + k, 0) = (node(k, 0) + node((k + 1) Mod 3, 0)) / 2#: node(3 + k, 1) = (node(k, 1) + node((k + 1) Mod 3, 1)) / 2#
        Next k
        a = ptri(p, 0): b = ptri(p, 1): cc = ptri(p, 2)
        det = (pxy(b, 0) - pxy(a, 0)) * (pxy(cc, 1) - pxy(a, 1)) - (pxy(cc, 0) - pxy(a, 0)) * (pxy(b, 1) - pxy(a, 1))
        For k = 0 To 5
            x = node(k, 0): y = node(k, 1)
            l(0) = ((pxy(b, 0) - x) * (pxy(cc, 1) - y) - (pxy(cc, 0) - x) * (pxy(b, 1) - y)) / det
            l(1) = ((pxy(cc, 0) - x) * (pxy(a, 1) - y) - (pxy(a, 0) - x) * (pxy(cc, 1) - y)) / det
            l(2) = 1# - l(0) - l(1)
            For m = 0 To 2
                nb(m) = l(m) * (2# * l(m) - 1#): nb(3 + m) = 4# * l(m) * l((m + 1) Mod 3)
            Next m
            For c = 0 To 1
                x = 0#
                For m = 0 To 5: x = x + nb(m) * pvel(12 * p + 2 * m + c): Next m
                child(12 * e + 2 * k + c) = x
            Next c
        Next k
    Next e
    RPX_ProlongUpperVelocity = True
End Function
