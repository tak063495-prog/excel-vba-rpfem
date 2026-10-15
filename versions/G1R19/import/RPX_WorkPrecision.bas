Attribute VB_Name = "RPX_WorkPrecision"
Option Explicit
' Work-only precision repair. The constitutive/conic equations and tolerances are unchanged.
Public Function RPX_ZeroCostUpperModel() As Boolean
    Dim e As Long, id As Long, j As Long
    If RR <> 0 Or RPX_LoadPhase <> "reference_total" Or RStressScale <= 0# Then Exit Function
    For e = 0 To re - 1
        If RRigidId(e) >= 0 Or RMaterials(RMatId(e)).cohesion <> 0# Then Exit Function
        For j = 0 To 1: If RBody0(e, j) <> 0# Then Exit Function
        Next j
    Next e
    For id = 0 To RNEdge - 1
        If REdges(id).element(1) < 0 Then
            For j = 0 To 1: If REdges(id).load0(j) <> 0# Then Exit Function
            Next j
        End If
    Next id
    RPX_ZeroCostUpperModel = True
End Function
Private Sub WorkTerm(ByRef acc As RPX_DD, ByVal coefficient As Double, ByVal value As Double)
    If coefficient = 0# Or value = 0# Then Exit Sub
    acc = HDAdd(acc, HDMul(HDD(coefficient), HDD(value)))
End Sub
Public Function RPX_PreciseUpperWork(ByRef field() As Double) As Double
    Dim acc As RPX_DD, e As Long, k As Long, j As Long, id As Long, weight As Double, coefficient As Double
    If LBound(field) <> 0 Or UBound(field) < 12 * re - 1 Then err.Raise 9, "RPX_WorkPrecision", "UPPER_WORK_DIMENSIONS"
    ' Rebuild from the current physical mesh/loads: never use another mesh's retained work row.
    For e = 0 To re - 1
        For k = 3 To 5
            For j = 0 To 1
                coefficient = RArea(e) / 3# * RBody1(e, j) / RStressScale
                WorkTerm acc, coefficient, field(12 * e + 2 * k + j)
            Next j
        Next k
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            If .element(1) < 0 Then
                For k = 0 To 2
                    weight = 1#: If k = 2 Then weight = 4#
                    For j = 0 To 1
                        coefficient = .length * weight / 6# * .load1(j) / RStressScale
                        WorkTerm acc, coefficient, field(12 * .element(0) + 2 * .loc(0, k) + j)
                    Next j
                Next k
            End If
        End With
    Next id
    RPX_PreciseUpperWork = HDValue(acc)
End Function
Public Function RPX_NormalizeZeroCostUpper(ByRef candidate As RPX_HsdCandidateState) As Boolean
    Dim w As Double, i As Long, j As Long, k As Long, p As Long, residual As Double, worst As Double
    Dim acc As RPX_DD, value As Double, norm As Double, coneError As Double
    If candidate.kind <> "OPTIMAL_POINT" Or candidate.mode <> "upper" Or Not candidate.numericalGate Then Exit Function
    If Not RPX_ZeroCostUpperModel() Then Exit Function
    ' The original objective and every cone offset must be EXACTLY zero.
    For i = 0 To XN - 1: If XQ(i) <> 0# Then Exit Function
    Next i
    For i = 0 To XB - 1
        For j = 0 To XC(i).dimn - 1: If XC(i).offset(j) <> 0# Then Exit Function
        Next j
    Next i
    Dim trial As RPX_HsdCandidateState
    trial = candidate
    w = RPX_PreciseUpperWork(trial.x)
    If Not RPX_SchurFinite(w) Or w <= 0# Then Exit Function
    If Abs(w - 1#) <= 0.0000001 Or Abs(w - 1#) > 0.00001 Then Exit Function
    For i = 0 To XN - 1: trial.x(i) = trial.x(i) / w: Next i
    trial.field = trial.x
    ' Re-evaluate the compiled ORIGINAL equations after normalization, using DD products.
    For i = 0 To XM - 1
        acc = HDD(-XERhs(i))
        For p = XEPtr(i) To XEPtr(i + 1) - 1: WorkTerm acc, XEVal(p), trial.x(XEId(p)): Next p
        residual = Abs(HDValue(acc)): If residual > worst Then worst = residual
    Next i
    If worst > 0.00000001 Then Exit Function
    For i = 0 To XB - 1
        norm = 0#
        For j = 0 To XC(i).dimn - 1
            acc = HDD(0#)
            For k = 0 To XC(i).count - 1: WorkTerm acc, XC(i).coef(j, k), trial.x(XC(i).ids(k)): Next k
            value = HDValue(acc): trial.slack(XC(i).first + j) = value
            If j > 0 Then norm = norm + value * value
        Next j
        residual = Sqr(norm) - trial.slack(XC(i).first)
        If residual > coneError Then coneError = residual
    Next i
    If coneError > 0.0000001 Then Exit Function
    ' Constant-zero objective: the zero dual is feasible with objective/gap exactly zero.
    ' This proof applies only to the homogeneous zero-cost problem, never a general HSD certificate.
    For i = LBound(trial.y) To UBound(trial.y): trial.y(i) = 0#: Next i
    For i = LBound(trial.dual) To UBound(trial.dual): trial.dual(i) = 0#: Next i
    trial.kind = "ZERO_COST_FEASIBLE_POINT": trial.objective = 0#: trial.gap = 0#
    trial.primalResidual = worst: trial.dualResidual = 0#
    RPX_DiagEvent "upper_work_normalize", "proof=ZERO_COST_FEASIBLE_POINT;W=" & w & ";eq=" & worst & ";cone=" & coneError & ";physical_audit=pending"
    candidate = trial
    RPX_NormalizeZeroCostUpper = True
End Function
