Attribute VB_Name = "IssueSolverProbe"
Option Explicit
Public IssueForceFactor As Boolean, IssueHSDCalls As Long
Public Function IssueEligible(ByVal number As Long, ByVal message As String, ByVal cancel As Boolean, ByVal audit As Boolean) As Boolean
    XCancel = cancel: RPX_AuditRejected = audit
    IssueEligible = RPX_RobustEligible(number, message)
    XCancel = False: RPX_AuditRejected = False
End Function
Public Function IssueSolve(ByVal scenario As String, ByVal prefix As String) As String
    Dim v As Double
    On Error GoTo Failed
    XCancel = False: Call RPX_InputEnd: Call RPX_TestDisarm: Call RPX_ReadInputs
    RPX_SlopeDefinition 16
    RMaterials(0).cohesion = 5#: RMaterials(0).friction = 45#: RMaterials(0).dilation = 45#
    RStressScale = 5#: RSubdivisions = 2: RFsTolerance = 0.0001
    Call RPX_GenerateMesh: RFsActive = True: Call RPX_TotalReferenceLoads
    IssueHSDCalls = 0: IssueForceFactor = (scenario = "factor")
    If scenario = "memory" Then RPX_TestArm "robust_legacy", 7, "OUT_OF_MEMORY"
    If scenario = "subscript" Then RPX_TestArm "robust_legacy", 9, "SUBSCRIPT"
    If scenario = "cancel" Then RPX_TestArm "robust_legacy", 18, "ANALYSIS_CANCELLED"
    If scenario = "unknown" Then RPX_TestArm "robust_legacy", 5, "UNKNOWN_FAILURE"
    If scenario = "audit" Then RPX_TestArm "robust_legacy", 5, "RAW_YIELD_AUDIT_FAILED"
    If scenario = "model" Then
        RPX_TestArm "robust_legacy", 5, "NEWTON_FACTORIZATION_FAILED pivot=123"
        RPX_TestQueue "robust_rescue", 5, "MUTATE_EDGE"
    End If
    If scenario = "hsd_failure" Then
        RPX_TestArm "robust_legacy", 5, "NEWTON_FACTORIZATION_FAILED pivot=123"
        RPX_TestQueue "issue_hsd", 7, "HSD_TEST_MEMORY"
    End If
    RPX_DiagStart "G1R12_issue2_" & scenario
    v = RPX_Bound("upper", 3#)
    RPX_ExportExchangeX prefix & ".bin": MeshExport prefix & "_mesh.bin"
    IssueSolve = "PASS;value=" & RValue & ";audit=" & RAudit & ";engine=" & RPX_RobustEngine & ";rescues=" & RPX_RobustRescues & ";hsd_calls=" & IssueHSDCalls & ";status=" & XStatus
    RPX_DiagFinish "TEST_PASS", IssueSolve: RFsActive = False
    Exit Function
Failed:
    IssueSolve = "ERROR " & err.number & ":" & err.description & ";engine=" & RPX_RobustEngine & ";rescues=" & RPX_RobustRescues & ";hsd_calls=" & IssueHSDCalls
    RPX_DiagFinish "TEST_ERROR", IssueSolve: RFsActive = False: Call RPX_TestDisarm
End Function
