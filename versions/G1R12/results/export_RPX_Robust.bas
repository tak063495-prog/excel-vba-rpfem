Option Explicit
' G1R7: numerical engine choice is independent of literal c=0 geometry policy.
Public RPX_RobustRescues As Long
Public RPX_RobustEngine As String
Private rejectedPoint As String, rejectedNumber As Long, rejectedSource As String, rejectedText As String
Private Function PointKey(ByVal model As String, ByVal mode As String) As String
    PointKey = mode & "|" & RPX_InputEpoch & "|" & RSubdivisions & "|" & RPX_LoadPhase & "|" & RPX_ExactDouble(RStrengthFactor) & "|" & model
End Function
' Accept only the detailed token emitted by FactorSystem; retain its text in evidence.
Private Function FactorizationFailure(ByVal message As String) As Boolean
    Const prefix As String = "NEWTON_FACTORIZATION_FAILED pivot="
    Dim digits As String, i As Long, first As Long, negative As Boolean, value As Double
    If Left$(message, Len(prefix)) <> prefix Then Exit Function
    digits = mid$(message, Len(prefix) + 1): first = 1
    If Left$(digits, 1) = "-" Then negative = True: first = 2
    If Len(digits) < first Or Len(digits) - first + 1 > 10 Then Exit Function
    For i = first To Len(digits)
        If mid$(digits, i, 1) < "0" Or mid$(digits, i, 1) > "9" Then Exit Function
    Next i
    value = CDbl(mid$(digits, first))
    If negative Then
        If value > 2147483648# Then Exit Function
    Else
        If value > 2147483647# Then Exit Function
    End If
    FactorizationFailure = True
End Function
Public Function RPX_RobustEligible(ByVal number As Long, ByVal message As String) As Boolean
    If XCancel Or RPX_AuditRejected Or number <> 5 Then Exit Function
    Select Case Trim$(message)
        Case "CONIC_STEP_STALLED", "CONIC_NOT_CONVERGED", "NEWTON_FACTORIZATION_FAILED", "NT_NONINTERIOR_POINT", "NT_CORRECTOR_SINGULAR"
            RPX_RobustEligible = True
        Case Else
            RPX_RobustEligible = FactorizationFailure(Trim$(message))
    End Select
End Function
Public Sub RPX_OptimizeBound(ByVal prepared As Boolean, ByVal mode As String, ByRef useHSD As Boolean)
    Dim model As String, phase As String, factor As String, epoch As Long, q As Long
    Dim errorNumber As Long, errorSource As String, errorText As String, errorLine As Long
    RPX_RobustEngine = "LEGACY": RPX_RobustRescues = 0
    RPX_HSDKind = "": RPX_HSDCandidateId = 0: Erase RPX_HSDProofX
    If useHSD Then
        RPX_RobustEngine = "HSD_C0"
        RPX_OptimizeHSD prepared, mode
        Exit Sub
    End If
    RPX_AssertInputs "robust_begin"
    model = RPX_ExactModel(): phase = RPX_LoadPhase
    factor = RPX_ExactDouble(RStrengthFactor): epoch = RPX_InputEpoch: q = RSubdivisions
    If rejectedPoint = PointKey(model, mode) Then
        RPX_RobustEngine = "REJECTED_POINT"
        RPX_DiagEvent "robust_repeat_rejected", "mode=" & mode & ";Fs=" & RStrengthFactor & ";reason=" & rejectedText & ";exact_identity=True"
        err.Raise rejectedNumber, rejectedSource, rejectedText
    End If
    ' Compile once before the error boundary; construction failures never rescue.
    If Not prepared Then
        RPX_ResetPreparedState
        RPX_Compile
    End If
    On Error GoTo LegacyFailed
    RPX_TestHook "robust_legacy"
    RPX_Optimize prepared, mode, True
    Exit Sub
LegacyFailed:
    errorNumber = err.number: errorSource = err.source: errorText = err.description: errorLine = Erl
    Resume LegacyRejected
LegacyRejected:
    On Error GoTo 0
    If Not RPX_RobustEligible(errorNumber, errorText) Then err.Raise errorNumber, errorSource, errorText
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, "robust_legacy"
    RPX_AssertInputs "robust_rescue"
    If epoch <> RPX_InputEpoch Or phase <> RPX_LoadPhase Or factor <> RPX_ExactDouble(RStrengthFactor) Or q <> RSubdivisions Then err.Raise 5, "RPX_Robust", "ROBUST_CONTEXT_MISMATCH"
    If model <> RPX_ExactModel() Then err.Raise 5, "RPX_Robust", "ROBUST_MODEL_MISMATCH"
    RPX_TestHook "robust_rescue"
    RPX_AssertInputs "robust_rescue_hook"
    If epoch <> RPX_InputEpoch Or phase <> RPX_LoadPhase Or factor <> RPX_ExactDouble(RStrengthFactor) Or q <> RSubdivisions Then err.Raise 5, "RPX_Robust", "ROBUST_CONTEXT_MISMATCH"
    If model <> RPX_ExactModel() Then err.Raise 5, "RPX_Robust", "ROBUST_MODEL_MISMATCH"
    ' Force a new numeric factor and cold primal/dual state on the same compiled problem.
Call RPX_SchurRelease:     RPX_HLURelease: RPX_ResetPreparedState
    RPX_RobustRescues = 1: RPX_RobustEngine = "HSD_RESCUE": useHSD = True
    RPX_DiagEvent "robust_rescue", "reason=" & errorText & ";mode=" & mode & ";Fs=" & RStrengthFactor & ";q=" & q & ";input_epoch=" & epoch & ";load_phase=" & phase & ";scope=ORIGINAL_CONIC;attempt=1"
    ' Only known numerical rejection is memoized. Every rescue failure propagates.
    On Error GoTo RescueFailed
    RPX_OptimizeHSD True, mode
    RPX_DiagEvent "robust_result", "engine=" & RPX_RobustEngine & ";status=" & XStatus & ";kind=" & RPX_HSDKind & ";scope=ORIGINAL_CONIC"
    Exit Sub
RescueFailed:
    errorNumber = err.number: errorSource = err.source: errorText = err.description
    Resume RescueRejected
RescueRejected:
    On Error GoTo 0
    If Not XCancel And Not RPX_AuditRejected And errorNumber = 5 Then
        Select Case Trim$(errorText)
            Case "HSD_LINEAR_QUALITY_FAILED", "HSD_BORDER_QUALITY_FAILED", "HSD_PIVOT_FACTORIZATION_FAILED", "HSD_STEP_STALLED", "HSD_NOT_CONVERGED", "HSD_SCHUR_SINGULAR"
                RPX_AssertInputs "robust_reject"
                If epoch = RPX_InputEpoch And phase = RPX_LoadPhase And factor = RPX_ExactDouble(RStrengthFactor) And q = RSubdivisions And model = RPX_ExactModel() Then
                    rejectedPoint = PointKey(model, mode): rejectedNumber = errorNumber: rejectedSource = errorSource: rejectedText = errorText
                End If
        End Select
    End If
    err.Raise errorNumber, errorSource, errorText
End Sub