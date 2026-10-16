Attribute VB_Name = "ReadModelFix3Probe"
Option Explicit

' The legacy text model format has no constrained-segment records.
Public Function RM3Read(ByVal modelPath As String, ByVal prior As String) As Variant
    Dim key As String, revision As Long, count As Long, missing As Boolean
    Dim number As Long, message As String
    Call RPX_InputEnd
    XCancel = False
    If prior = "fresh" Then
        Set RConstraints = Nothing
    ElseIf prior = "stale" Then
        Set RConstraints = RPX_Row()
        RConstraints("9998:9999") = Array(9998&, 9999&)
    End If
    On Error GoTo Failed
    Call RPX_ReadModel(modelPath)
    missing = (RConstraints Is Nothing)
    count = -1
    If Not missing Then count = RConstraints.count
    key = RPX_ExactModel()
    revision = RPX_ModelVersion()
    RM3Read = Array("RETURNED", 0&, "", missing, count, key, revision, rn, re)
    Exit Function
Failed:
    number = Err.Number: message = Err.Description
    missing = (RConstraints Is Nothing)
    count = -1
    If Not missing Then count = RConstraints.count
    RM3Read = Array("REJECTED", number, message, missing, count, "", 0&, rn, re)
End Function

Public Function RM3Solve(ByVal modelPath As String, ByVal outputPath As String, ByVal fresh As Boolean) As Variant
    Dim result As String
    Call RPX_InputEnd
    XCancel = False
    If fresh Then Set RConstraints = Nothing
    RRawAuditEnabled = True
    result = RPX_TestFs(modelPath, outputPath)
    RM3Solve = Array(result, RLowerFs, RUpperFs, RFsCalls, RSubdivisions, RAudit, RPhysicalGateAccepted, RRawLowerAudit(1), RRawLowerAudit(2), RRawLowerAudit(3), RRawLowerAudit(4), RFsSearchIncomplete, XStatus, RValue, XObjective, AccuracyDirectLastKind, AccuracyDirectLastAccepted)
End Function
