Attribute VB_Name = "BatchEvidence"
Option Explicit
' Read-only test evidence. No solver/state mutation and no failure injection.
Public Sub BatchStagePath(ByVal path As String)
    XLogPath = path
End Sub
Public Function BatchDefinition() As String
    On Error GoTo Failed
    Call RPX_ReadInputs
    BatchDefinition = DPCount & vbTab & DRCount & vbTab & DFCount & vbTab & RPX_UnionArea() & vbTab & RPX_AnalysisPolicy()
    Exit Function
Failed:
    BatchDefinition = "ERROR " & err.number & ":" & err.description
End Function
Private Sub BatchState(ByRef state As RPX_MeshState, ByVal path As String)
    Dim f As Integer, i As Long, j As Long, n As Long, value As Double, rec As Variant, key As Variant, ends As Variant, kind As Long
    Dim number As Long, message As String, source As String
    On Error GoTo Failed
    If Not state.complete Then err.Raise 5, "BatchEvidence", "STATE_INCOMPLETE"
    If Len(Dir$(path)) > 0 Then err.Raise 5, "BatchEvidence", "OUTPUT_EXISTS"
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_ExchangeMagic f, "RP25M001"
    Put #f, , state.nodes: Put #f, , state.elements: Put #f, , state.materialN: Put #f, , state.rigidN: Put #f, , state.q
    Put #f, , state.stressScale: Put #f, , state.factor: Put #f, , state.value: Put #f, , state.objective: Put #f, , state.normalizedAudit
    For i = 0 To state.nodes - 1
        For j = 0 To 1: value = state.xy(i, j): Put #f, , value: Next j
    Next i
    For i = 0 To state.elements - 1
        For j = 0 To 2: n = state.triangles(i, j): Put #f, , n: Next j
        n = state.materialId(i): Put #f, , n: n = state.rigidId(i): Put #f, , n
    Next i
    For i = 0 To state.materialN - 1
        With state.materials(i)
            Put #f, , .cohesion: Put #f, , .friction: Put #f, , .dilation: Put #f, , .gamma
        End With
    Next i
    For i = 0 To state.elements - 1
        For j = 0 To 1
            value = state.body0(i, j): Put #f, , value: value = state.body1(i, j): Put #f, , value
        Next j
    Next i
    n = state.boundary.count: Put #f, , n
    For Each key In state.boundary.keys
        ends = Split(CStr(key), ":"): n = CLng(ends(0)): Put #f, , n: n = CLng(ends(1)): Put #f, , n
        rec = state.boundary(key)
        Select Case rec(0)
            Case "load": kind = 0
            Case "fixed": kind = 1
            Case "roller_x": kind = 2
            Case "roller_y": kind = 3
            Case Else: err.Raise 5, "BatchEvidence", "UNSUPPORTED_BOUNDARY"
        End Select
        Put #f, , kind: n = rec(1): Put #f, , n
        For j = 2 To 5: value = rec(j): Put #f, , value: Next j
    Next key
    n = UBound(state.field) - LBound(state.field) + 1: Put #f, , n
    For i = LBound(state.field) To UBound(state.field): value = state.field(i): Put #f, , value: Next i
    Close #f: Exit Sub
Failed:
    number = err.number: message = err.description: source = err.source
    On Error Resume Next: Close #f: On Error GoTo 0
    err.Raise number, source, message
End Sub
Public Function BatchResult(ByVal prefix As String) As String
    On Error GoTo Failed
    BatchResult = RPX_LastFailureNumber & vbTab & RPX_LastError & vbTab & RPX_LastFailureSource & vbTab & RPX_ResultCurrent & vbTab & RFsSearchIncomplete & vbTab & RFsLargestRootBracket & vbTab & RFsSearchNote & vbTab & RPX_DiagPath & vbTab & RBestLower.q & vbTab & RBestUpper.q & vbTab & RBestLower.phase & vbTab & RBestUpper.phase & vbTab & RBestLower.lowerRootValid & vbTab & RBestUpper.upperRootValid & vbTab & RBestLower.searchIncomplete & vbTab & RBestUpper.searchIncomplete
    If RPX_LastFailureNumber <> 0 Or Not RPX_ResultCurrent Then Exit Function
    If RPX_C0Ran Then err.Raise 5, "BatchEvidence", "C0_AUTO_UNEXPECTED"
    BatchState RBestLower, prefix & "_lower.bin"
    BatchState RBestUpper, prefix & "_upper.bin"
    Exit Function
Failed:
    BatchResult = "ERROR " & err.number & ":" & err.description
End Function
