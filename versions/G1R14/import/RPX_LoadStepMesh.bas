Attribute VB_Name = "RPX_LoadStepMesh"
Option Explicit
' Repair only a two-triangle fan incompatible with a prescribed traction step.
' It adds local stress freedoms without moving any original node/boundary or changing a load.
Public Function RPX_RepairLoadSteps() As Boolean
    Dim incident() As Long, firstEdge() As Long, secondEdge() As Long, outerCount() As Long
    Dim marked() As Boolean, id As Long, node As Long, e As Long, f As Long, k As Long, other As Long
    Dim oldRE As Long, oldRN As Long, splitCount As Long, a As Long, b As Long, c As Long, center As Long
    Dim i As Long, j As Long, inner As Long, u As Double, v As Double, x As Double, y As Double, lengthProduct As Double
    Dim incompatible As Boolean, delta0 As Double, delta1 As Double
    If XCancel Then err.Raise 18, "RPX_LoadStepMesh", "ANALYSIS_CANCELLED"
    If RPX_LoadPhase <> "raw" Then err.Raise 5, "RPX_LoadStepMesh", "LOAD_STEP_REPAIR_PHASE"
    oldRE = re: oldRN = rn
    ReDim incident(0 To rn - 1): ReDim firstEdge(0 To rn - 1): ReDim secondEdge(0 To rn - 1): ReDim outerCount(0 To rn - 1)
    ReDim marked(0 To oldRE - 1)
    For e = 0 To re - 1
        For k = 0 To 2: node = RTri(e, k): incident(node) = incident(node) + 1: Next k
    Next e
    For id = 0 To RNEdge - 1
        If REdges(id).element(1) < 0 Then
            For k = 0 To 1
                node = REdges(id).a: If k = 1 Then node = REdges(id).b
                outerCount(node) = outerCount(node) + 1
                If outerCount(node) = 1 Then firstEdge(node) = id
                If outerCount(node) = 2 Then secondEdge(node) = id
            Next k
        End If
    Next id
    For node = 0 To oldRN - 1
        If incident(node) = 2 And outerCount(node) = 2 Then
            i = firstEdge(node): j = secondEdge(node): e = REdges(i).element(0): f = REdges(j).element(0)
            If REdges(i).kind = "load" And REdges(j).kind = "load" And e <> f And RRigidId(e) < 0 And RRigidId(f) < 0 Then
                a = REdges(i).a: If a = node Then a = REdges(i).b
                b = REdges(j).a: If b = node Then b = REdges(j).b
                u = RXY(a, 0) - RXY(node, 0): v = RXY(a, 1) - RXY(node, 1)
                x = RXY(b, 0) - RXY(node, 0): y = RXY(b, 1) - RXY(node, 1)
                lengthProduct = Sqr((u * u + v * v) * (x * x + y * y))
                If u * x + v * y < 0# And Abs(u * y - v * x) <= 0.000000000001 * lengthProduct Then
                    inner = -1
                    ' The third node of each incident triangle is their shared internal edge endpoint.
                    For k = 0 To 2
                        other = RTri(e, k)
                        If other <> node And other <> a Then inner = other
                    Next k
                    If inner >= 0 Then
                        x = RXY(inner, 0) - RXY(node, 0): y = RXY(inner, 1) - RXY(node, 1)
                        delta0 = (REdges(i).load0(0) - REdges(j).load0(0)) * y - (REdges(i).load0(1) - REdges(j).load0(1)) * x
                        delta1 = (REdges(i).load1(0) - REdges(j).load1(0)) * y - (REdges(i).load1(1) - REdges(j).load1(1)) * x
                        incompatible = (delta0 <> 0# Or delta1 <> 0#)
                        If incompatible Then marked(e) = True: marked(f) = True
                    End If
                End If
            End If
        End If
    Next node
    For e = 0 To oldRE - 1: If marked(e) Then splitCount = splitCount + 1
    Next e
    If splitCount = 0 Then Exit Function
    If rn + splitCount >= UBound(RXY, 1) - 3 Or re + 2 * splitCount - 1 > UBound(RTri, 1) Then err.Raise 7, "RPX_LoadStepMesh", "LOAD_STEP_REPAIR_CAPACITY"
    For e = 0 To oldRE - 1
        If marked(e) Then
            a = RTri(e, 0): b = RTri(e, 1): c = RTri(e, 2): center = rn
            RXY(rn, 0) = (RXY(a, 0) + RXY(b, 0) + RXY(c, 0)) / 3#
            RXY(rn, 1) = (RXY(a, 1) + RXY(b, 1) + RXY(c, 1)) / 3#: rn = rn + 1
            RPX_SetPositive e, a, b, center: RPX_SetPositive re, b, c, center: RPX_SetPositive re + 1, c, a, center: re = re + 2
        End If
    Next e
    RProgramKind = "": RPX_FsBracketReset
    RPX_DiagEvent "load_step_mesh_repair", "split_triangles=" & splitCount & ";nodes_before=" & oldRN & ";nodes_after=" & rn & ";elements_before=" & oldRE & ";elements_after=" & re & ";load_phase=raw"
    RPX_RepairLoadSteps = True
End Function
