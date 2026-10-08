Option Explicit
' Explicit conforming centroid split. Original edges/faces/materials/loads retained.
' Caller decides whether to evaluate a mesh candidate; this routine cannot certify Fs.
Public Function RPX_SplitFreeFaces(ByRef velocity() As Double, ByRef prolonged() As Double) As Long
    Dim surfaceNodes() As Boolean, marks() As Boolean, parents() As Long, nextXY() As Double, nextTri() As Long
    Dim nextMat() As Long, nextRigid() As Long, nextBody0() As Double, nextBody1() As Double
    Dim oldXY() As Double, oldTri() As Long, oldGrad() As Double, oldN As Long, oldE As Long
    Dim i As Long, e As Long, k As Long, j As Long, c As Long, count As Long, target As Long, node As Long
    Dim px As Double, py As Double, bary(0 To 2) As Double, weights(0 To 5) As Double, value As Double
    Dim area0 As Double, area1 As Double, number As Long, source As String, message As String
    RPX_AssertInputs "free_face_split_begin"
    If RR <> 0 Then Exit Function
    If UBound(velocity) < 12 * re - 1 Then err.Raise 9, "RPX_FreeFaceSplit", "SPLIT_VELOCITY_DIMENSION"
    oldN = rn: oldE = re: ReDim marks(0 To oldE - 1): ReDim surfaceNodes(0 To oldN - 1)
    For i = 0 To RNEdge - 1
        With REdges(i)
            If .element(1) < 0 And .kind = "load" And .load0(0) = 0# And .load0(1) = 0# And .load1(0) = 0# And .load1(1) = 0# Then surfaceNodes(.a) = True: surfaceNodes(.b) = True
        End With
    Next i
    For e = 0 To oldE - 1
        For k = 0 To 2: If surfaceNodes(RTri(e, k)) Then marks(e) = True
        Next k
        If marks(e) Then count = count + 1
        area0 = area0 + RArea(e)
    Next e
    If count = 0 Then Exit Function
    oldXY = RXY: oldTri = RTri: oldGrad = RGrad
    ReDim nextXY(0 To oldN + count - 1, 0 To 1): ReDim nextTri(0 To oldE + 2 * count - 1, 0 To 2)
    ReDim nextMat(0 To oldE + 2 * count - 1): ReDim nextRigid(0 To oldE + 2 * count - 1)
    ReDim nextBody0(0 To oldE + 2 * count - 1, 0 To 1): ReDim nextBody1(0 To oldE + 2 * count - 1, 0 To 1)
    ReDim parents(0 To oldE + 2 * count - 1): ReDim prolonged(0 To 12 * (oldE + 2 * count) - 1)
    For i = 0 To oldN - 1: For j = 0 To 1: nextXY(i, j) = oldXY(i, j): Next j: Next i
    target = 0: node = oldN
    For e = 0 To oldE - 1
        If e Mod 256 = 0 Then RPX_Stage "free face split " & e & "/" & oldE
        If marks(e) Then
            For j = 0 To 1
                For k = 0 To 2: nextXY(node, j) = nextXY(node, j) + oldXY(oldTri(e, k), j) / 3#: Next k
            Next j
            For k = 0 To 2
                nextTri(target, 0) = oldTri(e, k): nextTri(target, 1) = oldTri(e, (k + 1) Mod 3): nextTri(target, 2) = node
                parents(target) = e: target = target + 1
            Next k
            node = node + 1
        Else
            For k = 0 To 2: nextTri(target, k) = oldTri(e, k): Next k
            parents(target) = e: target = target + 1
        End If
    Next e
    For i = 0 To target - 1
        e = parents(i): nextMat(i) = RMatId(e): nextRigid(i) = RRigidId(e)
        For j = 0 To 1: nextBody0(i, j) = RBody0(e, j): nextBody1(i, j) = RBody1(e, j): Next j
        For k = 0 To 5
            If k < 3 Then
                px = nextXY(nextTri(i, k), 0): py = nextXY(nextTri(i, k), 1)
            Else
                px = (nextXY(nextTri(i, k - 3), 0) + nextXY(nextTri(i, (k - 2) Mod 3), 0)) / 2#
                py = (nextXY(nextTri(i, k - 3), 1) + nextXY(nextTri(i, (k - 2) Mod 3), 1)) / 2#
            End If
            For c = 0 To 1: bary(c) = oldGrad(e, c, 0) * (px - oldXY(oldTri(e, 2), 0)) + oldGrad(e, c, 1) * (py - oldXY(oldTri(e, 2), 1)): Next c
            bary(2) = 1# - bary(0) - bary(1)
            For c = 0 To 2: weights(c) = bary(c) * (2# * bary(c) - 1#): weights(c + 3) = 4# * bary(c) * bary((c + 1) Mod 3): Next c
            For j = 0 To 1
                value = velocity(12 * e + j)
                For c = 1 To 5: value = value + weights(c) * (velocity(12 * e + 2 * c + j) - velocity(12 * e + j)): Next c
                prolonged(12 * i + 2 * k + j) = value
            Next j
        Next k
    Next i
    ' No DoEvents during publication of complete arrays. Same loads, no transform.
    RXY = nextXY: RTri = nextTri: RMatId = nextMat: RRigidId = nextRigid: RBody0 = nextBody0: RBody1 = nextBody1
    rn = node: re = target: Call RPX_Topology
    For e = 0 To re - 1: area1 = area1 + RArea(e): Next e
    If Abs(area1 - area0) > 0.000000001 * (1# + area0) Then err.Raise 5, "RPX_FreeFaceSplit", "SPLIT_AREA_CHANGED"
    Call RPX_FsBracketReset: Call RPX_LowerCacheReset: Call RPX_ResetPreparedState: RProgramKind = ""
    RPX_SplitFreeFaces = count
End Function