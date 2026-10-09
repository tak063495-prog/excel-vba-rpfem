Attribute VB_Name = "RepairProbe"
Option Explicit
Public Function RepairOriginalControl(ByVal path As String, ByVal output As String) As String
    Dim f As Integer, n As Long, i As Long, value As Double, candidate As RPX_HsdCandidateState
    Dim original() As Double, w As Double, good As Boolean
    f = FreeFile: Open path For Binary Access Read As #f
    Seek #f, LOF(f) - 8# * XN - 3#: Get #f, , n
    If n <> XN Then err.Raise 5, "RepairProbe", "FIELD_LENGTH"
    ReDim original(0 To XN - 1)
    For i = 0 To XN - 1: Get #f, , value: original(i) = value: Next i
    Close #f
    w = RPX_PreciseUpperWork(original)
    candidate.kind = "OPTIMAL_POINT": candidate.mode = "upper": candidate.numericalGate = True
    candidate.x = original: ReDim candidate.slack(0 To XS - 1): ReDim candidate.dual(0 To XS - 1): ReDim candidate.y(0 To XM)
    good = RPX_NormalizeZeroCostUpper(candidate)
    If Not good Then RepairOriginalControl = "FAIL" & vbTab & w: Exit Function
    xx = candidate.x: XObjective = 0#: RPX_AuditUpper True
    RPX_ExportExchangeX output
    RepairOriginalControl = IIf(RPhysicalGateAccepted, "PASS", "FAIL") & vbTab & w & vbTab & RReferenceWork & vbTab & RAudit & vbTab & candidate.kind & vbTab & candidate.primalResidual & vbTab & RPX_DiagAuditSummary()
End Function
Public Function RepairFinalAudit(ByVal output As String) As String
    RPX_AuditUpper True: RPX_ExportExchangeX output
    RepairFinalAudit = RPhysicalGateAccepted & vbTab & RReferenceWork & vbTab & RAudit & vbTab & RPX_HSDKind & vbTab & RPX_DiagAuditSummary()
End Function
