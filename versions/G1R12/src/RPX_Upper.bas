Attribute VB_Name = "RPX_Upper"
Option Explicit
' version=20260926_0630 engine=20260926_edge_reduce
Public RWork As Object, RWork0 As Object
Private zeroPhiCosts As Collection
Private Sub AddWork(ByVal id As Long, ByVal load0 As Double, ByVal load1 As Double)
    RPX_Add RWork, id, load1 / RStressScale: RPX_Add RWork0, id, load0 / RStressScale
    RPX_Cost RPX_Unit(id, -load0 / RStressScale)
End Sub
Public Function RPX_RigidVelocity(ByVal rigidId As Long, ByVal px As Double, ByVal py As Double, ByVal component As Long) As Object
    Dim row As Object, start As Long
    start = 12 * re + 3 * rigidId: Set row = RPX_Unit(start + component)
    If component = 0 Then RPX_Add row, start + 2, -(py - RRigids(rigidId).center(1)) Else RPX_Add row, start + 2, px - RRigids(rigidId).center(0)
    Set RPX_RigidVelocity = row
End Function
Public Sub RPX_VelocityPosition(ByVal e As Long, ByVal k As Long, ByRef px As Double, ByRef py As Double)
    Dim a As Long, b As Long
    If k < 3 Then
        px = RXY(RTri(e, k), 0): py = RXY(RTri(e, k), 1)
    Else
        a = RTri(e, k - 3): b = RTri(e, (k - 2) Mod 3)
        px = (RXY(a, 0) + RXY(b, 0)) / 2#: py = (RXY(a, 1) + RXY(b, 1)) / 2#
    End If
End Sub
Private Sub StrainRows(ByVal e As Long, ByRef bary() As Double, ByRef trace As Object, ByRef dev As Object, ByRef shear As Object)
    Dim g(0 To 5, 0 To 1) As Double, k As Long, id As Long
    RPX_P2Grad e, bary, g
    Set trace = RPX_Row(): Set dev = RPX_Row(): Set shear = RPX_Row()
    For k = 0 To 5
        id = 12 * e + 2 * k
        RPX_Add trace, id, g(k, 0): RPX_Add trace, id + 1, g(k, 1)
        RPX_Add dev, id, g(k, 0): RPX_Add dev, id + 1, -g(k, 1)
        RPX_Add shear, id, g(k, 1): RPX_Add shear, id + 1, g(k, 0)
    Next k
End Sub
Private Function TraceRow(ByRef rows() As Object, ByVal s As Double) As Object
    Dim w(0 To 2) As Double, k As Long
    RPX_TraceWeights s, w: Set TraceRow = RPX_Row()
    For k = 0 To 2: RPX_Axpy TraceRow, rows(k), w(k): Next k
End Function
Private Function BernsteinRow(ByRef rows() As Object, ByVal a As Double, ByVal b As Double, ByVal k As Long) As Object
    If k = 0 Then
        Set BernsteinRow = TraceRow(rows, a)
    ElseIf k = 2 Then
        Set BernsteinRow = TraceRow(rows, b)
    Else
        Set BernsteinRow = RPX_Row()
        RPX_Axpy BernsteinRow, TraceRow(rows, (a + b) / 2#), 2#
        RPX_Axpy BernsteinRow, TraceRow(rows, a), -0.5: RPX_Axpy BernsteinRow, TraceRow(rows, b), -0.5
    End If
End Function
Public Sub RPX_AssembleUpper()
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim e As Long, f As Long, k As Long, j As Long, q As Long, i As Long, id As Long, rid As Long, n As Long, m As Long, rho As Long
    Dim c As Double, phi As Double, bary(0 To 2) As Double, px As Double, py As Double, weight As Double
    Dim trace As Object, dev As Object, shear As Object, row As Object, dxr As Object, dyr As Object
    Dim jt(0 To 2) As Object, jn(0 To 2) As Object, tang As Object, normal As Object, bond As Boolean
    Dim wa() As Double, aa As Double, bb As Double, multiplier As Double, kind As String
    Dim bw() As Double, w0(0 To 2) As Double, w1(0 To 2) As Double, wm(0 To 2) As Double
    Dim directIds(0 To 12) As Long, directCoef(0 To 12) As Double, component As Long, sign As Long, index As Long, direction As Double, side As Long
    Dim edgeBefore As Long, edgeAfter As Long
    RPX_Begin 12 * re + 3 * RR: Set RWork = RPX_Row(): Set RWork0 = RPX_Row(): Set zeroPhiCosts = New Collection: n = RSubdivisions
    If n < 1 Or n > 16 Then err.Raise 5, , "INVALID_QUADRATURE_SUBDIVISIONS"
    ReDim wa(0 To n, 0 To n)
    ReDim bw(0 To n - 1, 0 To 2, 0 To 2)
    For q = 0 To n - 1
        RPX_TraceWeights q / n, w0: RPX_TraceWeights (q + 1) / n, w1: RPX_TraceWeights (q + 0.5) / n, wm
        For j = 0 To 2
            bw(q, 0, j) = w0(j): bw(q, 2, j) = w1(j): bw(q, 1, j) = 2# * wm(j) - (w0(j) + w1(j)) / 2#
        Next j
    Next q
    For i = 0 To n - 1
        For j = 0 To n - i - 1
            wa(i, j) = wa(i, j) + 1#: wa(i + 1, j) = wa(i + 1, j) + 1#: wa(i, j + 1) = wa(i, j + 1) + 1#
            If i + j < n - 1 Then wa(i + 1, j) = wa(i + 1, j) + 1#: wa(i + 1, j + 1) = wa(i + 1, j + 1) + 1#: wa(i, j + 1) = wa(i, j + 1) + 1#
        Next j
    Next i
    For rid = 0 To RR - 1
        For j = 0 To 2
            id = 12 * re + 3 * rid + j
            If RRigids(rid).fixed(j) Then RPX_Equal RPX_Unit(id)
            AddWork id, RRigids(rid).load0(j), RRigids(rid).load1(j)
        Next j
    Next rid
    For e = 0 To re - 1
        If e Mod 256 = 0 Then RPX_Stage "upper elements " & e & "/" & re
        For k = 3 To 5
            For j = 0 To 1: AddWork 12 * e + 2 * k + j, RArea(e) / 3# * RBody0(e, j), RArea(e) / 3# * RBody1(e, j): Next j
        Next k
        rid = RRigidId(e)
        If rid >= 0 Then
            For k = 0 To 5
                RPX_VelocityPosition e, k, px, py
                For j = 0 To 1
                    Set row = RPX_Unit(12 * e + 2 * k + j): RPX_Axpy row, RPX_RigidVelocity(rid, px, py, j), -1#: RPX_Equal row
                Next j
            Next k
        Else
            RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
            For k = 0 To 2
                For j = 0 To 2: bary(j) = 0#: Next j
                bary(k) = 1#: StrainRows e, bary, trace, dev, shear
                If phi = 0# Then
                    RPX_Equal trace
                Else
                    ' Eliminate the local dilation epigraph analytically. This
                    ' preserves the cone and avoids thousands of noisy equalities.
                    Set row = RPX_Row(): RPX_Axpy row, trace, 1# / sinPhi
                    RPX_Cone3 row, dev, shear: RPX_Cost trace, RArea(e) * c / (3# * tanPhi)
                    XC(XB - 1).modelKind = 1: XC(XB - 1).owner = e
                End If
            Next k
            If phi = 0# Then
                For i = 0 To n
                    For j = 0 To n - i
                        bary(0) = 1# - (i + j) / n: bary(1) = i / n: bary(2) = j / n
                        StrainRows e, bary, trace, dev, shear
                        rho = RPX_Variable(RArea(e) * c * wa(i, j) / (3# * n * n))
                        zeroPhiCosts.Add rho
                        RPX_Cone3 RPX_Unit(rho), dev, shear
                    Next j
                Next i
            End If
        End If
    Next e
    For id = 0 To RNEdge - 1
        If id Mod 512 = 0 Then RPX_Stage "upper edges " & id & "/" & RNEdge
        With REdges(id)
            e = .element(0): f = .element(1)
            If f >= 0 Then
                If RRigidId(e) < 0 Or RRigidId(e) <> RRigidId(f) Then
                    bond = (RRigidId(e) >= 0 Or RRigidId(f) >= 0 Or RMatId(e) <> RMatId(f))
                    For k = 0 To 2
                        Set dxr = RPX_Unit(12 * f + 2 * .loc(1, k)): RPX_Add dxr, 12 * e + 2 * .loc(0, k), -1#
                        Set dyr = RPX_Unit(12 * f + 2 * .loc(1, k) + 1): RPX_Add dyr, 12 * e + 2 * .loc(0, k) + 1, -1#
                        If bond Then RPX_Equal dxr: RPX_Equal dyr
                        Set jt(k) = RPX_Row(): Set jn(k) = RPX_Row()
                        RPX_Axpy jt(k), dxr, .tangent(0): RPX_Axpy jt(k), dyr, .tangent(1)
                        RPX_Axpy jn(k), dxr, .normal(0): RPX_Axpy jn(k), dyr, .normal(1)
                    Next k
                    If Not bond Then
                        RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
                        If phi <> 0# Then edgeBefore = edgeBefore + 2 * (2 * n + 1)
                        For k = 0 To 2
                            If phi = 0# Then
                                RPX_Equal jn(k)
                            Else
                                weight = 1#: If k = 2 Then weight = 4#
                                RPX_Cost jn(k), weight * .length * c / (6# * tanPhi)
                            End If
                        Next k
                        For q = 0 To n - 1
                            For k = 0 To 2
                                If phi = 0# Or k = 1 Or (q = 0 And k = 0) Or (q = n - 1 And k = 2) Then
                                    If phi = 0# Then rho = RPX_Variable(.length * c / (3# * n)): zeroPhiCosts.Add rho
                                    For sign = -1 To 1 Step 2
                                        index = 0
                                        For side = 0 To 1
                                            For j = 0 To 2
                                                For component = 0 To 1
                                                    directIds(index) = 12 * .element(side) + 2 * .loc(side, j) + component
                                                    If phi = 0# Then direction = sign * .tangent(component) Else direction = .normal(component) + sign * tanPhi * .tangent(component)
                                                    directCoef(index) = (2 * side - 1) * bw(q, k, j) * direction: index = index + 1
                                                Next component
                                            Next j
                                        Next side
                                        If phi = 0# Then
                                            directIds(12) = rho: directCoef(12) = 1#: RPX_LinearDirect directIds, directCoef, 13
                                        Else
                                            RPX_LinearDirect directIds, directCoef, 12
                                            edgeAfter = edgeAfter + 1
                                            With XC(XB - 1)
                                                .modelKind = 2: .owner = id: .parameter1 = q: .parameter2 = k: .directionSign = sign
                                            End With
                                        End If
                                    Next sign
                                End If
                            Next k
                        Next q
                    End If
                End If
            Else
                kind = .kind
                For k = 0 To 2
                    weight = 1#: If k = 2 Then weight = 4#
                    m = .loc(0, k): RPX_VelocityPosition e, m, px, py
                    For j = 0 To 1
                        AddWork 12 * e + 2 * m + j, .length * weight / 6# * .load0(j), .length * weight / 6# * .load1(j)
                        Set row = RPX_Unit(12 * e + 2 * m + j)
                        If kind = "fixed" Or (kind = "roller_x" And j = 0) Or (kind = "roller_y" And j = 1) Then
                            RPX_Equal row
                        ElseIf kind = "rigid" Then
                            RPX_Axpy row, RPX_RigidVelocity(.rigidId, px, py, j), -1#: RPX_Equal row
                        End If
                    Next j
                Next k
            End If
        End With
    Next id
    If RWork.count = 0 Then err.Raise 5, , "ZERO_REFERENCE_LOAD"
    RPX_Equal RWork, 1#
    RPX_DiagEvent "edge_constraint_mode", "engine=20260926_edge_reduce;ver=20260926_0630;quadrature=" & n & ";scalar_before=" & edgeBefore & ";scalar_after=" & edgeAfter
End Sub

Public Sub RPX_RefreshUpper(ByVal oldFactor As Double)
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim i As Long, j As Long, e As Long, side As Long, k As Long, component As Long, q As Long, b As Long
    Dim c As Double, phi As Double, oldSin As Double, w0(0 To 2) As Double, w1(0 To 2) As Double, wm(0 To 2) As Double
    Dim bw(0 To 2) As Double, direction As Double, item As Variant
    For Each item In zeroPhiCosts: XQ(CLng(item)) = XQ(CLng(item)) * oldFactor / RStrengthFactor: Next item
    For i = 0 To XB - 1
        With XC(i)
            If .modelKind = 1 Then
                e = .owner
                Call RPX_StrengthTrigAt(e, oldFactor, c, phi, sinPhi, cosPhi, tanPhi)
                oldSin = sinPhi
                Call RPX_StrengthTrig(e, c, phi, sinPhi, cosPhi, tanPhi)
                For j = 0 To .count - 1: .coef(0, j) = .coef(0, j) * oldSin / sinPhi: Next j
                
                
                
            ElseIf .modelKind = 2 Then
                e = REdges(.owner).element(0): RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi: q = .parameter1: b = .parameter2
                RPX_TraceWeights q / RSubdivisions, w0: RPX_TraceWeights (q + 1) / RSubdivisions, w1: RPX_TraceWeights (q + 0.5) / RSubdivisions, wm
                For k = 0 To 2
                    If b = 0 Then
                        bw(k) = w0(k)
                    ElseIf b = 2 Then
                        bw(k) = w1(k)
                    Else
                        bw(k) = 2# * wm(k) - (w0(k) + w1(k)) / 2#
                    End If
                Next k
                For j = 0 To .count - 1
                    side = j \ 6: k = (j Mod 6) \ 2: component = j Mod 2
                    direction = REdges(.owner).normal(component) + .directionSign * tanPhi * REdges(.owner).tangent(component)
                    .coef(0, j) = (2 * side - 1) * bw(k) * direction
                Next j
            End If
        End With
    Next i
End Sub
