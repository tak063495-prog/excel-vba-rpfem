Attribute VB_Name = "RPX_AccuracyMesh"
Option Explicit
' Optional lower-mesh angular enrichment, equivalent to Python scope=star.
' Every old edge and vertex is retained; no constitutive or load change.
Public Function RPX_AngularFanMesh(ByVal elementCap As Long, ByRef parent() As Long) As Boolean
    Dim exposed() As Boolean, selected() As Boolean, childArea() As Double
    Dim xy() As Double, tri() As Long, mat() As Long, rigid() As Long
    Dim body0() As Double, body1() As Double, oldArea() As Double
    Dim oldRE As Long, oldRN As Long, splitN As Long, id As Long, e As Long, k As Long
    Dim a As Long, b As Long, c As Long, center As Long, nextE As Long, nextN As Long, child As Long
    RPX_AssertInputs "angular_fan_plan"
    If elementCap < re Or elementCap > 50000 Then err.Raise 5, "RPX_AccuracyMesh", "ANGULAR_FAN_ELEMENT_CAP"
    oldRE = re: oldRN = rn: oldArea = RArea
    ReDim exposed(0 To rn - 1): ReDim selected(0 To re - 1)
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0)
            If .element(1) < 0 And .kind = "load" And RRigidId(e) < 0 Then
                If RMaterials(RMatId(e)).cohesion = 0# Then
                    If Abs(.load0(0)) + Abs(.load0(1)) + Abs(.load1(0)) + Abs(.load1(1)) = 0# Then
                        exposed(.a) = True: exposed(.b) = True
                    End If
                End If
            End If
        End With
    Next id
    ' Include the entire star, including adjacent cohesive or rigid triangles.
    For e = 0 To oldRE - 1
        For k = 0 To 2
            If exposed(RTri(e, k)) Then selected(e) = True: splitN = splitN + 1: Exit For
        Next k
    Next e
    If splitN = 0 Then Exit Function
    If oldRE + 2 * splitN > elementCap Then
        RPX_DiagEvent "angular_fan_skip", "reason=element_cap;old_elements=" & oldRE & ";proposed_elements=" & (oldRE + 2 * splitN) & ";cap=" & elementCap
        Exit Function
    End If
    If oldRN + splitN > 100000 Then err.Raise 7, "RPX_AccuracyMesh", "ANGULAR_FAN_NODE_CAPACITY"
    ReDim xy(0 To 99999, 0 To 1): ReDim tri(0 To 219999, 0 To 2)
    ReDim mat(0 To oldRE + 2 * splitN - 1): ReDim rigid(0 To oldRE + 2 * splitN - 1)
    ReDim body0(0 To oldRE + 2 * splitN - 1, 0 To 1): ReDim body1(0 To oldRE + 2 * splitN - 1, 0 To 1)
    ReDim parent(0 To oldRE + 2 * splitN - 1)
    For e = 0 To oldRN - 1
        For k = 0 To 1: xy(e, k) = RXY(e, k): Next k
    Next e
    nextE = oldRE: nextN = oldRN
    For e = 0 To oldRE - 1
        If e Mod 256 = 0 Then RPX_Stage "angular fan preparation " & e
        a = RTri(e, 0): b = RTri(e, 1): c = RTri(e, 2)
        For k = 0 To 2: tri(e, k) = RTri(e, k): Next k
        parent(e) = e
        If selected(e) Then
            center = nextN: nextN = nextN + 1
            xy(center, 0) = (RXY(a, 0) + RXY(b, 0) + RXY(c, 0)) / 3#
            xy(center, 1) = (RXY(a, 1) + RXY(b, 1) + RXY(c, 1)) / 3#
            tri(e, 0) = a: tri(e, 1) = b: tri(e, 2) = center
            tri(nextE, 0) = b: tri(nextE, 1) = c: tri(nextE, 2) = center: parent(nextE) = e
            tri(nextE + 1, 0) = c: tri(nextE + 1, 1) = a: tri(nextE + 1, 2) = center: parent(nextE + 1) = e
            nextE = nextE + 2
        End If
    Next e
    For child = 0 To nextE - 1
        e = parent(child): mat(child) = RMatId(e): rigid(child) = RRigidId(e)
        For k = 0 To 1: body0(child, k) = RBody0(e, k): body1(child, k) = RBody1(e, k): Next k
    Next child
    RPX_AssertInputs "angular_fan_commit"
    RXY = xy: RTri = tri: RMatId = mat: RRigidId = rigid: RBody0 = body0: RBody1 = body1
    rn = nextN: re = nextE
    RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
    ' Topology reads the unchanged boundary dictionary in its current load phase.
    ' Reassigning regions here would discard the immutable parent load payload.
    Call RPX_Topology
    ReDim childArea(0 To oldRE - 1)
    For child = 0 To re - 1: childArea(parent(child)) = childArea(parent(child)) + RArea(child): Next child
    For e = 0 To oldRE - 1
        If Abs(childArea(e) - oldArea(e)) > 0.0000000001 * (1# + oldArea(e)) Then err.Raise 5, "RPX_AccuracyMesh", "ANGULAR_FAN_AREA_CHANGED"
    Next e
    RPX_DiagEvent "angular_fan_candidate", "scope=star;split_elements=" & splitN & ";old_elements=" & oldRE & ";new_elements=" & re & ";old_vertices_preserved=1;old_edges_preserved=1;parent_material_and_loads=1;certificate=0"
    RPX_AngularFanMesh = True
End Function

' Optional eight-triangle upper candidate for one supported homogeneous c0 slope.
' The six exterior vertices and original face definitions stay unchanged.
Public Function RPX_C0SlopeWedgeMesh(ByVal depth As Double) As Boolean
    Dim xy() As Double, tri() As Long, mat() As Long, rigid() As Long
    Dim body0() As Double, body1() As Double, row As Variant, rows As Variant
    Dim e As Long, k As Long, id As Long, boundaryCount As Long, det As Double, area As Double, px As Double, py As Double
    RPX_AssertInputs "c0_wedge_plan"
    If Not RPX_SchurFinite(depth) Then Exit Function
    If depth <= 0# Then Exit Function
    If RM <> 1 Or RR <> 0 Or DRCount <> 1 Or DPCount <> 6 Then Exit Function
    If RPX_WaterCouplingOn() Or RPX_SeismicKhSigned() <> 0# Then Exit Function
    If RMaterials(0).cohesion <> 0# Or RMaterials(0).friction <= 0# Or RMaterials(0).friction >= 89# Or RMaterials(0).gamma <= 0# Then Exit Function
    If DRegions(0).materialId <> 0 Or DRegions(0).rigidId >= 0 Or DRegions(0).pointCount <> 6 Then Exit Function
    For k = 0 To 5
        If DRegions(0).points(k) <> k Then Exit Function
        If Not RPX_SchurFinite(DPXY(k, 0)) Or Not RPX_SchurFinite(DPXY(k, 1)) Then Exit Function
    Next k
    If Not (DPXY(0, 0) < DPXY(4, 0) And DPXY(4, 0) < DPXY(3, 0) And DPXY(3, 0) < DPXY(1, 0)) Then Exit Function
    If Not (DPXY(0, 1) < DPXY(2, 1) And DPXY(2, 1) < DPXY(4, 1)) Then Exit Function
    If DPXY(1, 1) <> DPXY(0, 1) Or DPXY(2, 0) <> DPXY(1, 0) Or DPXY(3, 1) <> DPXY(2, 1) Then Exit Function
    If DPXY(5, 0) <> DPXY(0, 0) Or DPXY(5, 1) <> DPXY(4, 1) Or depth >= DPXY(4, 1) - DPXY(2, 1) Then Exit Function
    For id = 0 To DFCount - 1
        If DFaces(id).kind = "rigid" Or DFaces(id).rigidId >= 0 Then Exit Function
        For k = 0 To 1
            If DFaces(id).load0(k) <> 0# Or DFaces(id).load1(k) <> 0# Then Exit Function
        Next k
    Next id
    ReDim xy(0 To 99999, 0 To 1): ReDim tri(0 To 219999, 0 To 2)
    ReDim mat(0 To 7): ReDim rigid(0 To 7): ReDim body0(0 To 7, 0 To 1): ReDim body1(0 To 7, 0 To 1)
    For e = 0 To 5: For k = 0 To 1: xy(e, k) = DPXY(e, k): Next k: Next e
    xy(6, 0) = DPXY(4, 0): xy(6, 1) = DPXY(4, 1) - depth
    xy(7, 0) = DPXY(4, 0): xy(7, 1) = (DPXY(0, 1) + DPXY(2, 1)) / 2#
    rows = Array(Array(6, 3, 4), Array(0, 1, 7), Array(1, 2, 7), Array(2, 3, 7), Array(7, 3, 6), Array(7, 6, 5), Array(6, 4, 5), Array(0, 7, 5))
    For e = 0 To 7
        row = rows(e): For k = 0 To 2: tri(e, k) = CLng(row(k)): Next k
        det = (xy(tri(e, 1), 0) - xy(tri(e, 0), 0)) * (xy(tri(e, 2), 1) - xy(tri(e, 0), 1)) - (xy(tri(e, 2), 0) - xy(tri(e, 0), 0)) * (xy(tri(e, 1), 1) - xy(tri(e, 0), 1))
        If det <= 0.000000000001 Or Not RPX_SchurFinite(det) Then Exit Function
        area = area + det / 2#: px = 0#: py = 0#
        For k = 0 To 2: px = px + xy(tri(e, k), 0) / 3#: py = py + xy(tri(e, k), 1) / 3#: Next k
        If Not RPX_InRegion(px, py, 0) Then Exit Function
        mat(e) = 0: rigid(e) = -1
        If RGravityMode = "reference" Then body1(e, 1) = -RMaterials(0).gamma Else body0(e, 1) = -RMaterials(0).gamma
    Next e
    If Abs(area - DRegions(0).area) > 0.000000001 Then Exit Function
    RPX_AssertInputs "c0_wedge_commit"
    RXY = xy: RTri = tri: RMatId = mat: RRigidId = rigid: RBody0 = body0: RBody1 = body1: rn = 8: re = 8
    Set RBoundary = RPX_Row(): Set RConstraints = RPX_Row()
    For k = 0 To 5: RConstraints(RPX_EdgeKey(k, (k + 1) Mod 6)) = Array(k, (k + 1) Mod 6): Next k
    RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
    RPX_LoadPhase = "raw": Call RPX_Topology: Call RPX_MapFaces
    area = 0#: For e = 0 To re - 1: area = area + RArea(e): Next e
    For id = 0 To RNEdge - 1: If REdges(id).element(1) < 0 Then boundaryCount = boundaryCount + 1
    Next id
    If boundaryCount <> 6 Or Abs(area - DRegions(0).area) > 0.000000001 Then err.Raise 5, "RPX_AccuracyMesh", "C0_WEDGE_ORIGINAL_DOMAIN_CHANGED"
    RPX_DiagEvent "c0_slope_wedge_candidate", "elements=8;depth=" & depth & ";original_area=" & area & ";original_six_vertices_and_faces=1;homogeneous_c0_dry_gravity=1;certificate=0"
    RPX_C0SlopeWedgeMesh = True
End Function

' Analytical threshold is a search hint only, never a bound or certificate.
Public Function RPX_C0SlopeWedgeHint(ByVal depth As Double) As Double
    Dim low As Double, high As Double, mid As Double, c As Double, phi As Double, angle As Double, i As Long
    angle = Atn((DPXY(4, 1) - DPXY(3, 1) - depth) / (DPXY(3, 0) - DPXY(4, 0)))
    low = 0.0001: high = 1000#
    For i = 1 To 64
        mid = (low + high) / 2#: Call RPX_StrengthAt(0, mid, c, phi)
        If phi > angle Then low = mid Else high = mid
    Next i
    RPX_C0SlopeWedgeHint = (low + high) / 2#
End Function
