Option Explicit
' Factor-free candidate construction. A returned field is not an optimal root.
Public FactorFreeReason As String
Public Function RPX_Phi0QCandidate(ByRef velocity() As Double, ByVal factor As Double, ByVal newQ As Long, ByRef field() As Double, ByRef candidateFs As Double, ByRef candidateValue As Double) As Boolean
    Dim e As Long, i As Long, qOld As Long, fOld As Double, xOld() As Double, hasX As Boolean
    Dim audit As Double, raw() As Double, expected As Double, identity As String
    Dim number As Long, source As String, message As String
    RPX_AssertInputs "phi0_q_begin": FactorFreeReason = "INELIGIBLE"
    If RR <> 0 Or factor <= 0# Or RPX_LoadPhase <> "reference_total" Then Exit Function
    If newQ <> 2 And newQ <> 4 And newQ <> 8 Then err.Raise 5, "RPX_FactorFree", "PHI0_Q_ARGUMENT"
    If UBound(velocity) < 12 * re - 1 Then err.Raise 9, "RPX_FactorFree", "PHI0_VELOCITY_DIMENSION"
    For e = 0 To re - 1
        If RRigidId(e) >= 0 Or RMaterials(RMatId(e)).friction <> 0# Then Exit Function
        If RBody0(e, 0) <> 0# Or RBody0(e, 1) <> 0# Then Exit Function
    Next e
    For i = 0 To RNEdge - 1
        If REdges(i).load0(0) <> 0# Or REdges(i).load0(1) <> 0# Then Exit Function
    Next i
    qOld = RSubdivisions: fOld = RStrengthFactor: identity = RPX_ExactModel()
    hasX = VectorReady(xx): If hasX Then xOld = xx
    On Error GoTo Failed
    Call RPX_ResetPreparedState: RProgramKind = ""
    RSubdivisions = newQ: RStrengthFactor = factor: Call RPX_AssembleUpper
    ReDim field(0 To XN - 1)
    For i = 0 To 12 * re - 1: field(i) = velocity(i): Next i
    Call RPX_LiftUpperEpigraph(field)
    If Not RPX_OriginalPrimalGate(field) Then FactorFreeReason = "ORIGINAL_GATE": GoTo Done
    expected = OriginalCandidateObjective
    If Not RPX_FieldGate("upper", field, factor, newQ, expected, candidateValue, audit, raw) Then FactorFreeReason = "PHYSICAL_GATE": GoTo Done
    candidateFs = factor * candidateValue / 0.999999
    If candidateFs <= 0# Or Not RPX_SchurFinite(candidateFs) Then FactorFreeReason = "NONPOSITIVE_FS": GoTo Done
    RStrengthFactor = candidateFs: Call RPX_AssembleUpper
    ReDim field(0 To XN - 1)
    For i = 0 To 12 * re - 1: field(i) = velocity(i): Next i
    Call RPX_LiftUpperEpigraph(field)
    If Not RPX_OriginalPrimalGate(field) Then FactorFreeReason = "FINAL_ORIGINAL_GATE": GoTo Done
    expected = OriginalCandidateObjective
    If Not RPX_FieldGate("upper", field, candidateFs, newQ, expected, candidateValue, audit, raw) Then FactorFreeReason = "FINAL_PHYSICAL_GATE": GoTo Done
    If candidateValue >= 1# Then FactorFreeReason = "WORK_GATE": GoTo Done
    RPX_Phi0QCandidate = True: FactorFreeReason = "TARGET_WITNESS_ROOT_FALSE"
Done:
    RSubdivisions = qOld: RStrengthFactor = fOld
    RProgramKind = "": Call RPX_ResetPreparedState
    If hasX Then xx = xOld Else Erase xx
    If identity <> RPX_ExactModel() Then err.Raise 5, "RPX_FactorFree", "PHI0_IDENTITY_CHANGED"
    Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    RSubdivisions = qOld: RStrengthFactor = fOld: RProgramKind = "": Call RPX_ResetPreparedState
    If hasX Then xx = xOld Else Erase xx
    err.Raise number, source, message
End Function
Private Function VectorReady(ByRef values() As Double) As Boolean
    Dim n As Long
    On Error GoTo Missing
    n = UBound(values): VectorReady = n >= LBound(values): Exit Function
Missing:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
End Function

Public Function RPX_DirectC0Candidate(ByVal factor As Double, ByRef field() As Double, ByRef value As Double) As Boolean
    Dim e As Long, i As Long, j As Long, k As Long, side As Long, exterior() As Long, internal() As Long, slip() As Long
    Dim fOld As Double, qOld As Long, oldX() As Double, hasX As Boolean, identity As String
    Dim c As Double, phi As Double, sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim tx As Double, ty As Double, nx As Double, ny As Double, vx As Double, vy As Double, work As Double
    Dim audit As Double, raw() As Double, number As Long, source As String, message As String
    RPX_AssertInputs "direct_c0_begin": FactorFreeReason = "INELIGIBLE"
    If RM <> 1 Or RR <> 0 Or factor <= 0# Or RPX_LoadPhase <> "reference_total" Then Exit Function
    If RMaterials(0).cohesion <> 0# Or RMaterials(0).friction <= 0# Or RMaterials(0).gamma <= 0# Then Exit Function
    For e = 0 To re - 1
        If RMatId(e) <> 0 Or RRigidId(e) >= 0 Then Exit Function
        If RBody0(e, 0) <> 0# Or RBody0(e, 1) <> 0# Or RBody1(e, 0) <> 0# Or RBody1(e, 1) <> -RMaterials(0).gamma Then Exit Function
    Next e
    ReDim exterior(0 To re - 1): ReDim internal(0 To re - 1): ReDim slip(0 To re - 1)
    For i = 0 To RNEdge - 1
        With REdges(i)
            If .kind = "rigid" Or .load0(0) <> 0# Or .load0(1) <> 0# Or .load1(0) <> 0# Or .load1(1) <> 0# Then Exit Function
            If .element(1) < 0 Then
                If .kind = "load" Then exterior(.element(0)) = exterior(.element(0)) + 1
            Else
                For side = 0 To 1
                    e = .element(side): internal(e) = internal(e) + 1: slip(e) = i
                Next side
            End If
        End With
    Next i
    fOld = RStrengthFactor: qOld = RSubdivisions: identity = RPX_ExactModel()
    hasX = VectorReady(xx): If hasX Then oldX = xx
    On Error GoTo Failed
    RStrengthFactor = factor: Call RPX_ResetPreparedState: RProgramKind = "": Call RPX_AssembleUpper
    RPX_StrengthTrig 0, c, phi, sinPhi, cosPhi, tanPhi
    For e = 0 To re - 1
        If exterior(e) = 2 And internal(e) = 1 Then
            With REdges(slip(e))
                tx = .tangent(0): ty = .tangent(1)
                If ty > 0# Then tx = -tx: ty = -ty
                nx = .normal(0): ny = .normal(1)
                If .element(0) = e Then nx = -nx: ny = -ny
            End With
            vx = tx + tanPhi * nx: vy = ty + tanPhi * ny
            work = -RArea(e) * RMaterials(0).gamma * vy / RStressScale
            If work > 0# Then
                ReDim field(0 To XN - 1)
                For k = 0 To 5: field(12 * e + 2 * k) = vx / work: field(12 * e + 2 * k + 1) = vy / work: Next k
                If RPX_OriginalPrimalGate(field) Then
                    If RPX_FieldGate("upper", field, factor, qOld, 0#, value, audit, raw) Then
                        If value < 1# Then RPX_DirectC0Candidate = True: FactorFreeReason = "TARGET_WITNESS_ROOT_FALSE": Exit For
                    End If
                End If
            End If
        End If
    Next e
    If Not RPX_DirectC0Candidate Then FactorFreeReason = "NO_AUDITED_DIRECT_WEDGE"
    RStrengthFactor = fOld: RSubdivisions = qOld: RProgramKind = "": Call RPX_ResetPreparedState
    If hasX Then xx = oldX Else Erase xx
    If identity <> RPX_ExactModel() Then err.Raise 5, "RPX_FactorFree", "DIRECT_C0_IDENTITY_CHANGED"
    Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    RStrengthFactor = fOld: RSubdivisions = qOld: RProgramKind = "": Call RPX_ResetPreparedState
    If hasX Then xx = oldX Else Erase xx
    err.Raise number, source, message
End Function