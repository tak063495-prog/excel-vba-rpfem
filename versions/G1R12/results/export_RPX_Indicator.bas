Option Explicit
Public Function RPX_TestIndicator(ByVal inputPath As String) As String
    Dim upper As Double, lower As Double, velocity() As Double, stress() As Double, eta() As Double, total As Double, e As Long, factor As Double
    On Error GoTo Failed
    RPX_ReadModel inputPath: factor = RStrengthFactor
    upper = RPX_Bound("upper", factor): velocity = xx
    lower = RPX_Bound("lower", factor): stress = xx
    RPX_ComputeIndicator eta, velocity, stress
    For e = 0 To re - 1: total = total + eta(e): Next e
    If Abs(total - (upper - lower)) > 0.00001 Then err.Raise 5, , "GAP_INDICATOR_SUM_FAILED " & total & " vs " & upper - lower
    RPX_TestIndicator = "PASS sum=" & total & " gap=" & upper - lower: Exit Function
Failed:
    RPX_TestIndicator = "FAIL " & err.description
End Function
' Local dissipation minus stress power. Fields must use the same strength factor.
' This is a mesh indicator; acceptance uses independently audited global bounds.
Public Sub RPX_ComputeIndicator(ByRef eta() As Double, ByRef velocity() As Double, ByRef stress() As Double)
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim e As Long, f As Long, k As Long, j As Long, q As Long, id As Long, Slot As Long
    Dim bary(0 To 2) As Double, basis(0 To 5) As Double, g(0 To 5, 0 To 1) As Double
    Dim ex As Double, ey As Double, shear As Double, sxx As Double, syy As Double, tau As Double
    Dim c As Double, phi As Double, a As Double, weight As Double, dissip As Double, power As Double
    Dim weights(0 To 2) As Double, sw(0 To 2) As Double, s As Double, vx As Double, vy As Double, nx As Double, ny As Double
    Dim tx As Double, ty As Double, jumpNormal As Double, jumpTangent As Double, value As Double
    ReDim eta(0 To re - 1)
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
            For q = 0 To 5
                If q < 3 Then a = 0.445948490915965: weight = 0.223381589678011 Else a = 0.091576213509771: weight = 0.109951743655322
                For j = 0 To 2: bary(j) = a: Next j
                bary(q Mod 3) = 1# - 2# * a
                RPX_P2Grad e, bary, g
                For j = 0 To 2: basis(j) = bary(j) ^ 2: basis(j + 3) = 2# * bary(j) * bary((j + 1) Mod 3): Next j
                ex = 0#: ey = 0#: shear = 0#: sxx = 0#: syy = 0#: tau = 0#
                For k = 0 To 5
                    ex = ex + velocity(12 * e + 2 * k) * g(k, 0)
                    ey = ey + velocity(12 * e + 2 * k + 1) * g(k, 1)
                    shear = shear + velocity(12 * e + 2 * k) * g(k, 1) + velocity(12 * e + 2 * k + 1) * g(k, 0)
                    sxx = sxx + basis(k) * stress(RPX_StressId(e, k, 0))
                    syy = syy + basis(k) * stress(RPX_StressId(e, k, 1))
                    tau = tau + basis(k) * stress(RPX_StressId(e, k, 2))
                Next k
                If phi = 0# Then dissip = c * Sqr((ex - ey) ^ 2 + shear ^ 2) Else dissip = c * (ex + ey) / tanPhi
                power = sxx * ex + syy * ey + tau * shear
                eta(e) = eta(e) + RArea(e) * weight * (dissip - power)
            Next q
        End If
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1)
            If f >= 0 Then
                If RRigidId(e) < 0 And RRigidId(f) < 0 And RMatId(e) = RMatId(f) Then
                    RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi: value = 0#: nx = .normal(0): ny = .normal(1)
                    For q = 0 To 2
                        s = 0.5 + (q - 1) * Sqr(0.15): weight = 5# / 18#: If q = 1 Then weight = 4# / 9#
                        RPX_TraceWeights s, weights
                        sw(0) = (1# - s) ^ 2: sw(1) = s ^ 2: sw(2) = 2# * s * (1# - s)
                        vx = 0#: vy = 0#: tx = 0#: ty = 0#
                        For k = 0 To 2
                            vx = vx + weights(k) * (velocity(12 * f + 2 * .loc(1, k)) - velocity(12 * e + 2 * .loc(0, k)))
                            vy = vy + weights(k) * (velocity(12 * f + 2 * .loc(1, k) + 1) - velocity(12 * e + 2 * .loc(0, k) + 1))
                            Slot = .loc(0, k)
                            tx = tx + sw(k) * (stress(RPX_StressId(e, Slot, 0)) * nx + stress(RPX_StressId(e, Slot, 2)) * ny)
                            ty = ty + sw(k) * (stress(RPX_StressId(e, Slot, 2)) * nx + stress(RPX_StressId(e, Slot, 1)) * ny)
                        Next k
                        jumpNormal = vx * nx + vy * ny: jumpTangent = vx * .tangent(0) + vy * .tangent(1)
                        If phi = 0# Then dissip = c * Abs(jumpTangent) Else dissip = c * jumpNormal / tanPhi
                        value = value + .length * weight * (dissip - tx * vx - ty * vy)
                    Next q
                    eta(e) = eta(e) + value / 2#: eta(f) = eta(f) + value / 2#
                End If
            End If
        End With
    Next id
    For e = 0 To re - 1
        If eta(e) < 0# Then eta(e) = 0#
    Next e
End Sub