Option Explicit
Public Const RPX_HSDRepairVersion As String = "20260927_HSD_DD_OPERATOR_V2"
Public Type RPX_HLinearQuality
    valid As Boolean
    matrixEpoch As Long
    factorEpoch As Long
    rhsEpoch As Long
    corrections As Long
    checks As Long
    bestResidual As Double
    finalResidual As Double
    normwiseBackward As Double
    componentBackward As Double
    Residuals(0 To 8) As Double
End Type