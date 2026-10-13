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
