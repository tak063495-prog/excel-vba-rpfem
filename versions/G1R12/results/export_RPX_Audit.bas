Option Explicit
Public RValue As Double, RAudit As Double, RDissipation As Double, RReferenceWork As Double, RFixedWork As Double
Public RRawLowerAudit(1 To 4) As Double
Public RPhysicalGateAccepted As Boolean, RPhysicalGateReason As String
Public Function RPX_FieldGate(ByVal side As String, ByRef field() As Double, ByVal factor As Double, ByVal q As Long, ByVal expectedObjective As Double, ByRef measuredValue As Double, ByRef measuredAudit As Double, ByRef measuredRaw() As Double) As Boolean
    Dim oldX() As Double, oldFactor As Double, oldQ As Long, oldObjective As Double
    Dim oldValue As Double, oldAudit As Double, oldDissipation As Double, oldReference As Double, oldFixed As Double
    Dim oldRaw(1 To 4) As Double, oldGate As Boolean, oldReason As String
    Dim i As Long, number As Long, source As String, message As String, accepted As Boolean
    RPX_AssertInputs "field_gate"
    oldX = xx: oldFactor = RStrengthFactor: oldQ = RSubdivisions: oldObjective = XObjective
    oldValue = RValue: oldAudit = RAudit: oldDissipation = RDissipation: oldReference = RReferenceWork: oldFixed = RFixedWork
    oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
    For i = 1 To 4: oldRaw(i) = RRawLowerAudit(i): Next i
    On Error GoTo Failed
    RPX_TestHook "candidate_field_gate"
    xx = field: RStrengthFactor = factor: RSubdivisions = q: XObjective = expectedObjective
    Erase RRawLowerAudit
    If side = "lower" Then
        Call RPX_AuditLower(True)
    ElseIf side = "upper" Then
        Call RPX_AuditUpper(True)
    Else
        err.Raise 5, "RPX_Audit", "FIELD_GATE_SIDE"
    End If
    accepted = RPhysicalGateAccepted: measuredValue = RValue: measuredAudit = RAudit
    ReDim measuredRaw(1 To 4)
    For i = 1 To 4: measuredRaw(i) = RRawLowerAudit(i): Next i
Restore:
    xx = oldX: RStrengthFactor = oldFactor: RSubdivisions = oldQ: XObjective = oldObjective
    RValue = oldValue: RAudit = oldAudit: RDissipation = oldDissipation: RReferenceWork = oldReference: RFixedWork = oldFixed
    RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
    For i = 1 To 4: RRawLowerAudit(i) = oldRaw(i): Next i
    On Error GoTo 0
    If number <> 0 Then err.Raise number, source, message
    RPX_FieldGate = accepted: Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    Resume Restore
End Function
Private Sub CheckError(ByVal error As Double, ByVal category As Long, ByVal indexA As Long, Optional ByVal indexB As Long = -1, Optional ByVal indexC As Long = -1)
    If category >= 1 And category <= 4 Then
        If error * RStressScale > RRawLowerAudit(category) Then RRawLowerAudit(category) = error * RStressScale
    End If
    If error > RAudit Then RAudit = error
    RPX_DiagAuditCheck category, error, indexA, indexB, indexC
End Sub

Private Function StressTraction(ByVal e As Long, ByVal k As Long, ByVal j As Long, ByVal nx As Double, ByVal ny As Double) As Double
    If j = 0 Then
        StressTraction = xx(RPX_StressId(e, k, 0)) * nx + xx(RPX_StressId(e, k, 2)) * ny
    Else
        StressTraction = xx(RPX_StressId(e, k, 2)) * nx + xx(RPX_StressId(e, k, 1)) * ny
    End If
End Function
Private Sub AuditRigidLoad(ByRef externalLoad() As Double, ByVal rid As Long, ByVal px As Double, ByVal py As Double, ByVal fx As Double, ByVal fy As Double)
    externalLoad(rid, 0) = externalLoad(rid, 0) + fx / RStressScale: externalLoad(rid, 1) = externalLoad(rid, 1) + fy / RStressScale
    externalLoad(rid, 2) = externalLoad(rid, 2) + ((px - RRigids(rid).center(0)) * fy - (py - RRigids(rid).center(1)) * fx) / RStressScale
End Sub
Private Sub RPX_AuditLowerCore(ByVal candidateCheck As Boolean)
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim e As Long, f As Long, k As Long, j As Long, l As Long, id As Long, side As Long, rid As Long
    Dim c As Double, phi As Double, sxx As Double, syy As Double, tau As Double, residual As Double, bary(0 To 2) As Double
    Dim g(0 To 5, 0 To 1) As Double, force() As Double, externalLoad() As Double, traction(0 To 2, 0 To 1) As Double
    Dim nx As Double, ny As Double, px As Double, py As Double, w As Double, mx As Double, my As Double, rec As Variant
    Erase RRawLowerAudit
    RValue = xx(RLoadVariable): RAudit = 0#: ReDim force(0 To RR, 0 To 2): ReDim externalLoad(0 To RR, 0 To 2)
    For rid = 0 To RR - 1
        For j = 0 To 2: externalLoad(rid, j) = (RRigids(rid).load0(j) + RValue * RRigids(rid).load1(j)) / RStressScale: Next j
    Next rid
    For e = 0 To re - 1
        rid = RRigidId(e)
        If rid >= 0 Then
            px = 0#: py = 0#
            For k = 0 To 2: px = px + RXY(RTri(e, k), 0) / 3#: py = py + RXY(RTri(e, k), 1) / 3#: Next k
            AuditRigidLoad externalLoad, rid, px, py, RArea(e) * (RBody0(e, 0) + RValue * RBody1(e, 0)), RArea(e) * (RBody0(e, 1) + RValue * RBody1(e, 1))
        Else
            RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
            For k = 0 To 5
                sxx = xx(RPX_StressId(e, k, 0)): syy = xx(RPX_StressId(e, k, 1)): tau = xx(RPX_StressId(e, k, 2))
                CheckError Sqr((sxx - syy) ^ 2 + 4# * tau * tau) + sinPhi * (sxx + syy) - 2# * c * cosPhi, 1, e, k
            Next k
            For l = 0 To 2
                For k = 0 To 2: bary(k) = 0#: Next k
                bary(l) = 1#: RPX_P2Grad e, bary, g, True
                For j = 0 To 1
                    residual = (RBody0(e, j) + RValue * RBody1(e, j)) / RStressScale
                    For k = 0 To 5
                        residual = residual + xx(RPX_StressId(e, k, j)) * g(k, j) + xx(RPX_StressId(e, k, 2)) * g(k, 1 - j)
                    Next k
                    CheckError Abs(residual), 2, e, l, j
                Next j
            Next l
        End If
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1): side = 0: nx = .normal(0): ny = .normal(1)
            If RRigidId(e) >= 0 And f >= 0 Then
                If RRigidId(f) < 0 Then e = f: f = .element(0): side = 1: nx = -nx: ny = -ny
            End If
            rid = -1
            If RRigidId(e) >= 0 Then
                If f < 0 Then
                    rid = RRigidId(e): px = (RXY(.a, 0) + RXY(.b, 0)) / 2#: py = (RXY(.a, 1) + RXY(.b, 1)) / 2#
                    AuditRigidLoad externalLoad, rid, px, py, .length * (.load0(0) + RValue * .load1(0)), .length * (.load0(1) + RValue * .load1(1))
                End If
            Else
                For k = 0 To 2
                    For j = 0 To 1: traction(k, j) = StressTraction(e, .loc(side, k), j, nx, ny): Next j
                Next k
                If f >= 0 Then
                    rid = RRigidId(f)
                    If rid < 0 Then
                        For k = 0 To 2
                            For j = 0 To 1: CheckError Abs(traction(k, j) - StressTraction(f, .loc(1 - side, k), j, nx, ny)), 3, id, k: Next j
                        Next k
                    End If
                ElseIf .kind = "rigid" Then
                    rid = .rigidId: px = (RXY(.a, 0) + RXY(.b, 0)) / 2#: py = (RXY(.a, 1) + RXY(.b, 1)) / 2#
                    AuditRigidLoad externalLoad, rid, px, py, .length * (.load0(0) + RValue * .load1(0)), .length * (.load0(1) + RValue * .load1(1))
                Else
                    For k = 0 To 2
                        For j = 0 To 1
                            If .kind = "load" Or (.kind = "roller_x" And j = 1) Or (.kind = "roller_y" And j = 0) Then CheckError Abs(traction(k, j) - (.load0(j) + RValue * .load1(j)) / RStressScale), 4, id, j
                        Next j
                    Next k
                End If
                If rid >= 0 Then
                    For k = 0 To 2
                        If k = 0 Then
                            w = 1# / 12#
                        ElseIf k = 1 Then
                            w = 1# / 4#
                        Else
                            w = 1# / 6#
                        End If
                        For j = 0 To 1: force(rid, j) = force(rid, j) + .length / 3# * traction(k, j): Next j
                        mx = .length * ((RXY(.a, 0) - RRigids(rid).center(0)) / 3# + (RXY(.b, 0) - RXY(.a, 0)) * w)
                        my = .length * ((RXY(.a, 1) - RRigids(rid).center(1)) / 3# + (RXY(.b, 1) - RXY(.a, 1)) * w)
                        force(rid, 2) = force(rid, 2) + mx * traction(k, 1) - my * traction(k, 0)
                    Next k
                End If
            End If
        End With
    Next id
    For Each rec In RReactions
        force(rec(1), rec(2)) = force(rec(1), rec(2)) + rec(3) * xx(rec(0))
        force(rec(1), 2) = force(rec(1), 2) + rec(4) * xx(rec(0))
    Next rec
    For rid = 0 To RR - 1
        For j = 0 To 2
            If Not RRigids(rid).fixed(j) Then CheckError Abs(force(rid, j) - externalLoad(rid, j)), 5 + Abs(CLng(j = 2)), rid, j
        Next j
    Next rid
    If RAudit > 0.0000001 Then
        RPhysicalGateAccepted = False: RPhysicalGateReason = "LOWER_PHYSICAL_AUDIT_FAILED " & RAudit
        If Not candidateCheck Then err.Raise 5, "RPX_Audit", RPhysicalGateReason
    End If
    If RRawAuditEnabled Then
        For j = 1 To 4
            If RRawLowerAudit(j) > 0.0000001 Then
                RPhysicalGateAccepted = False: RPhysicalGateReason = "LOWER_RAW_AUDIT_FAILED category=" & j & ";raw=" & RRawLowerAudit(j)
                If Not candidateCheck Then err.Raise 5, "RPX_Audit", RPhysicalGateReason
            End If
        Next j
    End If
    RPX_DiagEvent "lower_raw_audit", "yield=" & RRawLowerAudit(1) & ";equilibrium=" & RRawLowerAudit(2) & ";traction=" & RRawLowerAudit(3) & ";boundary=" & RRawLowerAudit(4) & ";enabled=" & RRawAuditEnabled & ";rigid_force_moment=EXISTING_NORMALIZED_GATE"
End Sub
Private Function RigidValue(ByVal rid As Long, ByVal px As Double, ByVal py As Double, ByVal j As Long) As Double
    Dim start As Long
    start = 12 * re + 3 * rid: RigidValue = xx(start + j)
    If j = 0 Then RigidValue = RigidValue - (py - RRigids(rid).center(1)) * xx(start + 2) Else RigidValue = RigidValue + (px - RRigids(rid).center(0)) * xx(start + 2)
End Function
Private Sub VelocityGradient(ByVal e As Long, ByRef bary() As Double, ByRef trace As Double, ByRef dev As Double, ByRef shear As Double)
    Dim g(0 To 5, 0 To 1) As Double, k As Long, ex As Double, ey As Double
    RPX_P2Grad e, bary, g: ex = 0#: ey = 0#: shear = 0#
    For k = 0 To 5
        ex = ex + xx(12 * e + 2 * k) * g(k, 0): ey = ey + xx(12 * e + 2 * k + 1) * g(k, 1)
        shear = shear + xx(12 * e + 2 * k) * g(k, 1) + xx(12 * e + 2 * k + 1) * g(k, 0)
    Next k
    trace = ex + ey: dev = ex - ey
End Sub
Private Function TraceValue(ByRef values() As Double, ByVal s As Double) As Double
    Dim w(0 To 2) As Double, k As Long
    RPX_TraceWeights s, w
    For k = 0 To 2: TraceValue = TraceValue + w(k) * values(k): Next k
End Function
Private Sub RPX_AuditUpperCore(ByVal candidateCheck As Boolean)
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim e As Long, f As Long, k As Long, j As Long, i As Long, q As Long, id As Long, rid As Long, n As Long, loc As Long
    Dim c As Double, phi As Double, bary(0 To 2) As Double, trace As Double, dev As Double, shear As Double
    Dim wa() As Double, jt(0 To 2) As Double, jn(0 To 2) As Double, tc(0 To 2) As Double, nc(0 To 2) As Double
    Dim px As Double, py As Double, vx As Double, vy As Double, weight As Double, a As Double, b As Double, bond As Boolean
    RAudit = 0#: RDissipation = 0#: RReferenceWork = 0#: RFixedWork = 0#: n = RSubdivisions
    ReDim wa(0 To n, 0 To n)
    For i = 0 To n - 1
        For j = 0 To n - i - 1
            wa(i, j) = wa(i, j) + 1#: wa(i + 1, j) = wa(i + 1, j) + 1#: wa(i, j + 1) = wa(i, j + 1) + 1#
            If i + j < n - 1 Then wa(i + 1, j) = wa(i + 1, j) + 1#: wa(i + 1, j + 1) = wa(i + 1, j + 1) + 1#: wa(i, j + 1) = wa(i, j + 1) + 1#
        Next j
    Next i
    For rid = 0 To RR - 1
        For j = 0 To 2
            vx = xx(12 * re + 3 * rid + j)
            If RRigids(rid).fixed(j) Then CheckError Abs(vx), 11, rid, j
            RReferenceWork = RReferenceWork + vx * RRigids(rid).load1(j) / RStressScale
            RFixedWork = RFixedWork + vx * RRigids(rid).load0(j) / RStressScale
        Next j
    Next rid
    For e = 0 To re - 1
        For k = 3 To 5
            For j = 0 To 1
                vx = xx(12 * e + 2 * k + j) * RArea(e) / (3# * RStressScale)
                RReferenceWork = RReferenceWork + vx * RBody1(e, j): RFixedWork = RFixedWork + vx * RBody0(e, j)
            Next j
        Next k
        rid = RRigidId(e)
        If rid >= 0 Then
            For k = 0 To 5
                RPX_VelocityPosition e, k, px, py
                For j = 0 To 1: CheckError Abs(xx(12 * e + 2 * k + j) - RigidValue(rid, px, py, j)), 11, e, rid: Next j
            Next k
        Else
            RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
            For k = 0 To 2
                For j = 0 To 2: bary(j) = 0#: Next j
                bary(k) = 1#: VelocityGradient e, bary, trace, dev, shear
                If phi = 0# Then
                    CheckError Abs(trace), 7, e, k
                Else
                    CheckError sinPhi * Sqr(dev * dev + shear * shear) - trace, 7, e, k
                    RDissipation = RDissipation + RArea(e) * c / (3# * tanPhi) * trace
                End If
            Next k
            If phi = 0# Then
                For i = 0 To n
                    For j = 0 To n - i
                        bary(0) = 1# - (i + j) / n: bary(1) = i / n: bary(2) = j / n
                        VelocityGradient e, bary, trace, dev, shear
                        RDissipation = RDissipation + RArea(e) * c * wa(i, j) * Sqr(dev * dev + shear * shear) / (3# * n * n)
                    Next j
                Next i
            End If
        End If
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1)
            If f >= 0 Then
                bond = (RMatId(e) <> RMatId(f) Or RRigidId(e) >= 0 Or RRigidId(f) >= 0)
                For k = 0 To 2
                    vx = xx(12 * f + 2 * .loc(1, k)) - xx(12 * e + 2 * .loc(0, k))
                    vy = xx(12 * f + 2 * .loc(1, k) + 1) - xx(12 * e + 2 * .loc(0, k) + 1)
                    If bond Then CheckError Abs(vx), 8, id, 0: CheckError Abs(vy), 8, id, 1
                    jt(k) = vx * .tangent(0) + vy * .tangent(1): jn(k) = vx * .normal(0) + vy * .normal(1)
                Next k
                If Not bond Then
                    RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
                    For q = 0 To n - 1
                        a = q / n: b = (q + 1) / n
                        tc(0) = TraceValue(jt, a): tc(2) = TraceValue(jt, b): tc(1) = 2# * TraceValue(jt, (a + b) / 2#) - (tc(0) + tc(2)) / 2#
                        nc(0) = TraceValue(jn, a): nc(2) = TraceValue(jn, b): nc(1) = 2# * TraceValue(jn, (a + b) / 2#) - (nc(0) + nc(2)) / 2#
                        For k = 0 To 2
                            If phi = 0# Then
                                CheckError Abs(nc(k)), 9, id, q, k: RDissipation = RDissipation + .length * c * Abs(tc(k)) / (3# * n)
                            Else
                                CheckError tanPhi * Abs(tc(k)) - nc(k), 9, id, q, k
                            End If
                        Next k
                    Next q
                    If phi <> 0# Then RDissipation = RDissipation + .length * c / tanPhi * (jn(0) + jn(1) + 4# * jn(2)) / 6#
                End If
            Else
                For k = 0 To 2
                    loc = .loc(0, k): RPX_VelocityPosition e, loc, px, py: weight = 1#: If k = 2 Then weight = 4#
                    For j = 0 To 1
                        vx = xx(12 * e + 2 * loc + j)
                        RReferenceWork = RReferenceWork + vx * .load1(j) * weight * .length / (6# * RStressScale)
                        RFixedWork = RFixedWork + vx * .load0(j) * weight * .length / (6# * RStressScale)
                        If .kind = "fixed" Or (.kind = "roller_x" And j = 0) Or (.kind = "roller_y" And j = 1) Then CheckError Abs(vx), 10, id, j
                        If .kind = "rigid" Then CheckError Abs(vx - RigidValue(.rigidId, px, py, j)), 11, id, .rigidId
                    Next j
                Next k
            End If
        End With
    Next id
    CheckError Abs(RReferenceWork - 1#), 12, -1: RValue = RDissipation - RFixedWork
    RPX_DiagAuditCheck 13, Abs(RValue - XObjective), -1
    If RAudit > 0.0000001 Or Abs(RValue - XObjective) > 0.000002 Then
        RPhysicalGateAccepted = False: RPhysicalGateReason = "UPPER_PHYSICAL_AUDIT_FAILED " & RAudit
        If Not candidateCheck Then err.Raise 5, "RPX_Audit", RPhysicalGateReason
    End If
End Sub

Public Sub RPX_AuditLower(Optional ByVal candidateCheck As Boolean = False)
    Dim errorNumber As Long, errorText As String, errorSource As String, errorLine As Long
    On Error GoTo Failed
    RPX_DiagAuditStart "lower"
    RPhysicalGateAccepted = True: RPhysicalGateReason = ""
    Call RPX_AuditLowerCore(candidateCheck)
    If RPhysicalGateAccepted Then RPX_DiagAuditEnd "PASS" Else RPX_DiagAuditEnd "CANDIDATE_REJECTED"
    Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    RPX_DiagAuditEnd "FAIL"
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, "physical_audit"
    err.Raise errorNumber, errorSource, errorText
End Sub

Public Sub RPX_AuditUpper(Optional ByVal candidateCheck As Boolean = False)
    Dim errorNumber As Long, errorText As String, errorSource As String, errorLine As Long
    On Error GoTo Failed
    RPX_DiagAuditStart "upper"
    RPhysicalGateAccepted = True: RPhysicalGateReason = ""
    Call RPX_AuditUpperCore(candidateCheck)
    If RPhysicalGateAccepted Then RPX_DiagAuditEnd "PASS" Else RPX_DiagAuditEnd "CANDIDATE_REJECTED"
    Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    RPX_DiagAuditEnd "FAIL"
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, "physical_audit"
    err.Raise errorNumber, errorSource, errorText
End Sub