Attribute VB_Name = "WorkGateControls"
Option Explicit
Private original() As Double
Public Sub WorkGateRead(ByVal path As String)
    Dim f As Integer, i As Long, n As Long, value As Double
    f = FreeFile: Open path For Binary Access Read As #f
    Seek #f, LOF(f) - 8# * XN - 3#: Get #f, , n
    If n <> XN Then err.Raise 5, "WorkGateControls", "FIELD_LENGTH"
    ReDim original(0 To XN - 1)
    For i = 0 To XN - 1: Get #f, , value: original(i) = value: Next i
    Close #f
End Sub
Public Function WorkGateTest(ByVal kind As String) As String
    Dim c As RPX_HsdCandidateState, saved() As Double, i As Long, changed As Boolean, expected As Boolean
    Dim oldC As Double, oldFixed As Double, oldQ As Double, oldRR As Long, oldPhase As String
    Dim failure As Long, message As String, same As Boolean
    oldC = RMaterials(0).cohesion: oldFixed = RBody0(0, 1): oldQ = XQ(0): oldRR = RR: oldPhase = RPX_LoadPhase
    On Error GoTo Failed
    c.kind = "OPTIMAL_POINT": c.mode = "upper": c.numericalGate = True: c.x = original
    ReDim c.slack(0 To XS - 1): ReDim c.dual(0 To XS - 1): ReDim c.y(0 To XM)
    Select Case kind
        Case "original_repair": expected = True
        Case "zero": For i = 0 To XN - 1: c.x(i) = 0#: Next i
        Case "negative": For i = 0 To XN - 1: c.x(i) = -c.x(i): Next i
        Case "large_work": For i = 0 To XN - 1: c.x(i) = 2# * c.x(i): Next i
        Case "numerical_gate": c.numericalGate = False
        Case "wrong_mode": c.mode = "lower"
        Case "certificate": c.kind = "PRIMAL_INFEASIBILITY_CERT"
        Case "positive_c": RMaterials(0).cohesion = 1#
        Case "fixed_load": RBody0(0, 1) = -1#
        Case "positive_cost": XQ(0) = 1#
        Case "rigid": RR = 1
        Case "raw_phase": RPX_LoadPhase = "raw"
        Case "broken_equation"
            For i = 0 To XM - 1
                If XERhs(i) = 0# And XEPtr(i) < XEPtr(i + 1) Then
                    c.x(XEId(XEPtr(i))) = c.x(XEId(XEPtr(i))) + 1000# / XEVal(XEPtr(i)): Exit For
                End If
            Next i
        Case "field_dimensions": ReDim c.x(0 To 0)
        Case "slack_dimensions": ReDim c.slack(0 To 0)
        Case Else: err.Raise 5, "WorkGateControls", "UNKNOWN_CONTROL"
    End Select
    saved = c.x
    changed = RPX_NormalizeZeroCostUpper(c)
    If changed <> expected Then err.Raise 5, "WorkGateControls", "WRONG_GATE_DECISION"
    If Not changed Then
        same = (UBound(saved) = UBound(c.x))
        For i = 0 To UBound(saved)
            If RPX_ExactDouble(saved(i)) <> RPX_ExactDouble(c.x(i)) Then same = False
        Next i
        If Not same Then err.Raise 5, "WorkGateControls", "REJECTED_CANDIDATE_MUTATED"
    Else
        If c.kind <> "ZERO_COST_FEASIBLE_POINT" Or c.objective <> 0# Or c.gap <> 0# Then err.Raise 5, "WorkGateControls", "WRONG_ZERO_COST_PROOF"
        For i = 0 To UBound(c.y): If c.y(i) <> 0# Then err.Raise 5, "WorkGateControls", "NONZERO_DUAL"
        Next i
        For i = 0 To UBound(c.dual): If c.dual(i) <> 0# Then err.Raise 5, "WorkGateControls", "NONZERO_DUAL"
        Next i
        xx = c.x: XObjective = c.objective: RPX_AuditUpper True
        If Not RPhysicalGateAccepted Then err.Raise 5, "WorkGateControls", "REPAIRED_PHYSICAL_AUDIT_FAILED"
    End If
    WorkGateTest = "PASS" & vbTab & kind & vbTab & changed
    GoTo Restore
Failed:
    failure = err.number: message = err.description
    If failure = 9 Then
        For i = 0 To UBound(saved)
            If RPX_ExactDouble(saved(i)) <> RPX_ExactDouble(c.x(i)) Then failure = 5: message = "FAILED_CANDIDATE_MUTATED"
        Next i
    End If
    WorkGateTest = "ERROR" & vbTab & kind & vbTab & failure & vbTab & message
Restore:
    On Error GoTo 0
    RMaterials(0).cohesion = oldC: RBody0(0, 1) = oldFixed: XQ(0) = oldQ: RR = oldRR: RPX_LoadPhase = oldPhase
End Function
