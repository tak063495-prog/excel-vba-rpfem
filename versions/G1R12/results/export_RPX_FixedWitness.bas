Option Explicit
' Explicit experimental physical entry point, not an Fs root/optimizer replacement.
' Requires an already constructed mesh and the original reference-total load phase.
Public FixedWitnessStatus As String
Public Function RPX_TryFixedLowerWitness(ByVal factor As Double, Optional ByVal seconds As Double = 120#) As Boolean
    Dim oldX() As Double, restored() As Double, oldFactor As Double, oldValue As Double, oldObjective As Double
    Dim oldAudit As Double, oldQ As Long, modelKey As String, oldKind As String, oldFs As Double
    Dim oldLower() As Double, oldLowerFs As Double, hadX As Boolean, hadLower As Boolean, lowerAssigned As Boolean
    Dim number As Long, source As String, message As String, accepted As Boolean
    Dim measuredValue As Double, measuredAudit As Double, raw() As Double
    RPX_AssertInputs "fixed_lower_begin"
    If RPX_LoadPhase <> "reference_total" Then err.Raise 5, "RPX_FixedWitness", "FIXED_LOWER_REFERENCE_TOTAL_REQUIRED"
    If factor <= 0# Or seconds <= 0# Then err.Raise 5, "RPX_FixedWitness", "FIXED_LOWER_ARGUMENT"
    If AFActive Then err.Raise 5, "RPX_FixedWitness", "FIXED_LOWER_NESTED_CONTEXT"
    modelKey = RPX_ExactModel(): oldFactor = RStrengthFactor: oldQ = RSubdivisions
    oldLowerFs = RLowerFs: hadX = HasVector(xx): hadLower = HasVector(RLowerField)
    If hadX Then oldX = xx
    If hadLower Then oldLower = RLowerField
    oldValue = RValue: oldAudit = RAudit: oldObjective = XObjective: oldKind = RProgramKind: oldFs = RProgramFs
    On Error GoTo Failed
    Call RPX_ResetPreparedState: Call RPX_HLURelease
    RStrengthFactor = factor: Call RPX_AssembleLower
    Call RPX_FixLowerUnitLoad
    AFPhysicalGate = True: AFRawScale = RStressScale
    Call RPX_CondenseFixedWitness: Call RPX_Compile
    accepted = RPX_AHOFixedWitness(seconds, 80): FixedWitnessStatus = AHOStatus
    If accepted Then
        Call RPX_AffineRestoreX(xx, restored)
        If modelKey <> RPX_ExactModel() Then err.Raise 5, "RPX_FixedWitness", "FIXED_LOWER_IDENTITY_CHANGED"
        If Not RPX_FieldGate("lower", restored, factor, oldQ, 0#, measuredValue, measuredAudit, raw) Then accepted = False
        If Abs(measuredValue - 1#) > 0.0000001 Then accepted = False
    End If
    Call RPX_AffineEnd
    RStrengthFactor = oldFactor: RSubdivisions = oldQ
    Call RPX_AssembleLower
    RProgramKind = "": Call RPX_ResetPreparedState
    If accepted Then
        lowerAssigned = True: RLowerField = restored: RLowerFs = factor: xx = restored: RValue = measuredValue: RAudit = measuredAudit
        Dim i As Long
        For i = 1 To 4: RRawLowerAudit(i) = raw(i): Next i
        XStatus = "FIXED_FS_WITNESS"
        Call RPX_CertCapture("lower", "FIXED_FS_WITNESS", oldQ, False, "independent_AHO3_original_physical_gate")
        RPX_TryFixedLowerWitness = RPX_CertCurrent("lower") And RPX_CertLower.acceptedFs = factor
        If Not RPX_TryFixedLowerWitness Then
            If hadLower Then RLowerField = oldLower Else Erase RLowerField
            RLowerFs = oldLowerFs
            If hadX Then xx = oldX Else Erase xx
            RValue = oldValue: RAudit = oldAudit: XObjective = oldObjective
        End If
    Else
        If hadX Then xx = oldX Else Erase xx
        RValue = oldValue: RAudit = oldAudit: XObjective = oldObjective
    End If
    Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    Call RPX_AffineEnd: RStrengthFactor = oldFactor: RSubdivisions = oldQ
    RProgramKind = "": Call RPX_ResetPreparedState
    If hadX Then xx = oldX Else Erase xx
    If lowerAssigned Then
        If hadLower Then RLowerField = oldLower Else Erase RLowerField
        RLowerFs = oldLowerFs
    End If
    RValue = oldValue: RAudit = oldAudit: XObjective = oldObjective
    err.Raise number, source, message
End Function

Private Function HasVector(ByRef values() As Double) As Boolean
    Dim n As Long
    On Error GoTo Missing
    n = UBound(values): HasVector = (n >= LBound(values)): Exit Function
Missing:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
End Function